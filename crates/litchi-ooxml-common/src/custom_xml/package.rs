//! OPC package graph discovery and mutation for Custom XML Data Storage.

use crate::mce::Name;
use crate::signature::{refuse_signed_change, settle_unsigned_change};
use crate::{Error, Result};

use super::limit::LimitKind;
use litchi_opc::{OpcError, OpcPackage, PackURI, Part};
use std::collections::HashMap;
use std::sync::Arc;

use super::codec::{
    invalid, is_data_relationship, is_props_relationship, limit, read_props, require_at_most,
    validate_content_type, validate_payload,
};
use super::errors::{unsupported_content, unsupported_target};
use super::model::{Item, MAX_ITEMS, PROPS_CONTENT_TYPE, Props, Relationship};
use super::snapshot::{PartState, Snapshot};
use super::validation;

/// Discover and validate every explicit Custom XML Data Storage relationship.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn discover(package: &OpcPackage) -> Result<Vec<Item>> {
    let source_limits = package.read_limits();
    let mut items = Vec::new();
    scan(
        package,
        &source_limits,
        |source,
         source_relationship,
         part,
         data,
         root,
         props_part,
         props,
         props_xml,
         relationships| {
            items.push(Item::new(
                source.clone(),
                source_relationship.id.clone(),
                source_relationship,
                part,
                data.content_type().into(),
                root,
                data.blob_arc(),
                props_part,
                props,
                props_xml,
                relationships,
            ));
            Ok(())
        },
    )?;
    items.sort_unstable_by(|left, right| {
        left.source()
            .as_str()
            .cmp(right.source().as_str())
            .then_with(|| left.rel_id().cmp(right.rel_id()))
    });
    Ok(items)
}

fn scan(
    package: &OpcPackage,
    source_limits: &litchi_opc::SourceLimits,
    mut visit: impl FnMut(
        &PackURI,
        Relationship,
        PackURI,
        &dyn Part,
        Name,
        Option<PackURI>,
        Option<Props>,
        Option<Arc<Vec<u8>>>,
        Arc<[Relationship]>,
    ) -> Result<()>,
) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| is_data_relationship(relationship.reltype()))
    {
        return invalid("package root cannot source a Custom XML Data Storage relationship");
    }

    let mut occurrences = 0usize;
    let mut property_ids: HashMap<String, PackURI> = HashMap::new();
    let mut cached_props: HashMap<PackURI, Props> = HashMap::new();
    let mut props_owners: HashMap<PackURI, PackURI> = HashMap::new();
    let mut cached_roots: HashMap<PackURI, Name> = HashMap::new();

    for source in package.iter_parts() {
        for relationship in source
            .rels()
            .iter()
            .filter(|relationship| is_data_relationship(relationship.reltype()))
        {
            occurrences = occurrences
                .checked_add(1)
                .ok_or_else(|| limit(LimitKind::Items, MAX_ITEMS, usize::MAX))?;
            require_at_most(LimitKind::Items, occurrences, MAX_ITEMS)?;
            if relationship.is_external() {
                return Err(Error::Relationship(format!(
                    "custom XML relationship '{}' from '{}' must be internal",
                    relationship.r_id(),
                    source.partname().as_str()
                )));
            }
            let requested_name = relationship.target_partname().map_err(|error| {
                Error::Relationship(format!(
                    "invalid custom XML target '{}': {error}",
                    relationship.r_id()
                ))
            })?;
            let data = match package.get_part(&requested_name) {
                Ok(data) => data,
                Err(OpcError::PartNotFound(error)) => {
                    return Err(Error::Missing(format!(
                        "custom XML part '{}': {error}",
                        requested_name.as_str()
                    )));
                },
                Err(error) => return Err(error.into()),
            };
            let part = data.partname().clone();
            validate_content_type(data.content_type())?;
            let root = if let Some(root) = cached_roots.get(&part) {
                root.clone()
            } else {
                let root = validate_payload(data.blob(), source_limits)?;
                cached_roots.insert(part.clone(), root.clone());
                root
            };
            let relationships = relationships(data);
            let (props_part, props, props_xml) = resolve_props(
                package,
                data,
                &mut property_ids,
                &mut cached_props,
                &mut props_owners,
                source_limits,
            )?;
            visit(
                source.partname(),
                Relationship::from_opc(relationship),
                part,
                data,
                root,
                props_part,
                props,
                props_xml,
                relationships,
            )?;
        }
    }
    Ok(())
}

fn resolve_props(
    package: &OpcPackage,
    data: &dyn Part,
    property_ids: &mut HashMap<String, PackURI>,
    cache: &mut HashMap<PackURI, Props>,
    owners: &mut HashMap<PackURI, PackURI>,
    source_limits: &litchi_opc::SourceLimits,
) -> Result<(Option<PackURI>, Option<Props>, Option<Arc<Vec<u8>>>)> {
    let mut props_relationship = None;
    for relationship in data
        .rels()
        .iter()
        .filter(|relationship| is_props_relationship(relationship.reltype()))
    {
        if props_relationship.replace(relationship).is_some() {
            return invalid(format!(
                "custom XML part '{}' has more than one properties relationship",
                data.partname().as_str()
            ));
        }
    }
    let Some(relationship) = props_relationship else {
        return Ok((None, None, None));
    };
    if relationship.is_external() {
        return Err(Error::Relationship(
            "custom XML properties relationship must be internal".into(),
        ));
    }
    let requested_name = relationship.target_partname().map_err(|error| {
        Error::Relationship(format!("invalid custom XML properties target: {error}"))
    })?;
    let part = match package.get_part(&requested_name) {
        Ok(part) => part,
        Err(OpcError::PartNotFound(error)) => {
            return Err(Error::Missing(format!(
                "custom XML properties part '{}': {error}",
                requested_name.as_str()
            )));
        },
        Err(error) => return Err(error.into()),
    };
    let part_name = part.partname().clone();
    if let Some(existing_owner) = owners.insert(part_name.clone(), data.partname().clone())
        && existing_owner != *data.partname()
    {
        return invalid(format!(
            "custom XML properties part '{}' is shared by '{}' and '{}'",
            part_name.as_str(),
            existing_owner.as_str(),
            data.partname().as_str()
        ));
    }
    let props = if let Some(props) = cache.get(&part_name) {
        props.clone()
    } else {
        if part.content_type() != PROPS_CONTENT_TYPE {
            return Err(Error::ContentType {
                expected: PROPS_CONTENT_TYPE.into(),
                actual: part.content_type().into(),
            });
        }
        let props = read_props(part.blob(), source_limits)?;
        let key = props.id.to_ascii_lowercase();
        if let Some(existing) = property_ids.insert(key, part_name.clone())
            && existing != part_name
        {
            return invalid(format!("duplicate custom XML itemID '{}'", props.id));
        }
        cache.insert(part_name.clone(), props.clone());
        props
    };
    Ok((Some(part_name), Some(props), Some(part.blob_arc())))
}

fn relationships(part: &dyn Part) -> Arc<[Relationship]> {
    let mut values = part
        .rels()
        .iter()
        .map(Relationship::from_opc)
        .collect::<Vec<_>>();
    values.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    values.into()
}

/// Publish a complete desired Custom XML occurrence set onto an already
/// source-checked package clone.
pub(crate) fn apply_items(
    package: &mut OpcPackage,
    before: &Snapshot,
    desired: &[Item],
) -> Result<()> {
    let source_limits = package.read_limits();
    validation::items(desired, &source_limits)?;
    reject_part_aliases(desired)?;

    let mut staged = Vec::new();
    let mut source_names = Vec::new();
    for item in before.items() {
        push_name(&mut source_names, item.source());
    }
    for item in desired {
        push_name(&mut source_names, item.source());
    }

    for source_name in &source_names {
        let current = match package.get_part(source_name) {
            Ok(part) => part,
            Err(OpcError::PartNotFound(_)) => return Err(super::snapshot::source_mismatch()),
            Err(error) => return Err(error.into()),
        };
        let mut state = PartState::capture(current);
        let removed_ids = before
            .items()
            .iter()
            .filter(|item| item.source() == source_name)
            .map(Item::rel_id)
            .collect::<Vec<_>>();
        state
            .relationships
            .retain(|relationship| !removed_ids.iter().any(|id| *id == relationship.id));
        for item in desired.iter().filter(|item| item.source() == source_name) {
            if state
                .relationships
                .iter()
                .any(|relationship| relationship.id == item.source_relationship().id)
            {
                return Err(unsupported_target(
                    "custom_xml.apply",
                    "the Custom XML relationship target is already occupied",
                ));
            }
            state.relationships.push(item.source_relationship().clone());
        }
        state
            .relationships
            .sort_unstable_by(|left, right| left.id.cmp(&right.id));
        staged_state(&mut staged, state)?;
    }

    for item in desired {
        staged_state(
            &mut staged,
            PartState {
                name: item.part().clone(),
                content_type: item.content_type().into(),
                data: Arc::new(item.xml().to_vec()),
                relationships: item.relationships().to_vec(),
            },
        )?;

        if let (Some(props_part), Some(props_xml)) = (item.props_part(), item.props_xml()) {
            let mut state = before.source().part(props_part).cloned();
            if state.is_none() {
                state = match package.get_part(props_part) {
                    Ok(part) => Some(PartState::capture(part)),
                    Err(OpcError::PartNotFound(_)) => None,
                    Err(error) => return Err(error.into()),
                };
            }
            let mut state = state.unwrap_or_else(|| PartState {
                name: props_part.clone(),
                content_type: PROPS_CONTENT_TYPE.into(),
                data: Arc::new(Vec::new()),
                relationships: Vec::new(),
            });
            state.content_type = PROPS_CONTENT_TYPE.into();
            state.data = Arc::new(props_xml.to_vec());
            staged_state(&mut staged, state)?;
        }
    }

    validate_destinations(package, before, desired, &staged)?;
    let parts = staged
        .iter()
        .map(PartState::to_part)
        .collect::<Result<Vec<_>>>()
        .map_err(|error| match error {
            Error::Opc(OpcError::DuplicateRelationshipId(_)) => {
                Error::Internal("custom XML prepared part relationships are invalid")
            },
            error => error,
        })?;
    let staged_names = staged
        .iter()
        .map(|part| part.name.clone())
        .collect::<Vec<_>>();

    refuse_signed_change(package)?;
    for part in parts {
        package.add_part(Box::new(part));
    }

    let mut obsolete = before
        .source()
        .parts
        .iter()
        .filter(|part| !has_name(&staged_names, &part.name))
        .map(|part| part.name.clone())
        .collect::<Vec<_>>();
    obsolete.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    for name in obsolete {
        if !part_is_referenced(package, &name)? {
            package.remove_part(&name);
        }
    }
    settle_unsigned_change(package)?;
    Ok(())
}

fn reject_part_aliases(items: &[Item]) -> Result<()> {
    for item in items {
        if item.source() == item.part() {
            return Err(unsupported_target(
                "custom_xml.apply",
                "the Custom XML source cannot also be its data part",
            ));
        }
        if let Some(props_part) = item.props_part()
            && (props_part == item.part() || props_part == item.source())
        {
            return Err(unsupported_target(
                "custom_xml.apply",
                "the Custom XML properties part aliases an owning part",
            ));
        }
    }
    Ok(())
}

fn staged_state(staged: &mut Vec<PartState>, state: PartState) -> Result<()> {
    if let Some(existing) = staged.iter_mut().find(|part| part.name == state.name) {
        if *existing != state {
            return Err(unsupported_content(
                "custom_xml.apply",
                "the Custom XML graph has a shared part that cannot be edited independently",
            ));
        }
    } else {
        staged.push(state);
    }
    Ok(())
}

fn validate_destinations(
    package: &OpcPackage,
    before: &Snapshot,
    desired: &[Item],
    staged: &[PartState],
) -> Result<()> {
    for part in staged {
        let belongs_to_source = before
            .source()
            .parts
            .iter()
            .any(|existing| existing.name == part.name);
        let is_host = desired.iter().any(|item| item.source() == &part.name);
        match package.get_part(&part.name) {
            Ok(existing) => {
                if !belongs_to_source && !is_host {
                    return Err(unsupported_target(
                        "custom_xml.apply",
                        "the Custom XML destination part is already occupied",
                    ));
                }
                if belongs_to_source
                    && !before
                        .source()
                        .part(existing.partname())
                        .is_some_and(|source| source.matches(existing))
                {
                    return Err(super::snapshot::source_mismatch());
                }
            },
            Err(OpcError::PartNotFound(_)) if !belongs_to_source => {
                validate_new_part_name(package, &part.name)?;
            },
            Err(OpcError::PartNotFound(_)) => {
                return Err(super::snapshot::source_mismatch());
            },
            Err(error) => return Err(error.into()),
        }
    }
    Ok(())
}

fn part_is_referenced(package: &OpcPackage, target: &PackURI) -> Result<bool> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        if relationship.target_partname()? == *target {
            return Ok(true);
        }
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            if relationship.target_partname()? == *target {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn validate_new_part_name(package: &OpcPackage, name: &PackURI) -> Result<()> {
    match package.validate_new_part_name(name) {
        Ok(()) => Ok(()),
        Err(
            OpcError::DuplicatePartName(_)
            | OpcError::EquivalentPartNames { .. }
            | OpcError::DerivedPartNames { .. },
        ) => Err(unsupported_target(
            "custom_xml.apply",
            "the Custom XML destination part is already occupied",
        )),
        Err(error) => Err(error.into()),
    }
}

fn push_name(names: &mut Vec<PackURI>, name: &PackURI) {
    if !has_name(names, name) {
        names.push(name.clone());
    }
}

fn has_name(names: &[PackURI], candidate: &PackURI) -> bool {
    names
        .iter()
        .any(|name| name.as_str().eq_ignore_ascii_case(candidate.as_str()))
}
