//! OPC package graph discovery and mutation for Custom XML Data Storage.

use crate::mce::Name;
use crate::{Error, Result};
use litchi_opc::part::XmlPart;
use litchi_opc::{OpcPackage, PackURI, Part, TargetMode};
use std::collections::HashMap;
use std::sync::Arc;

use super::codec::{
    invalid, is_data_relationship, is_props_relationship, limit, read_props, require_at_most,
    require_rel_id, validate_content_type, validate_payload, validate_props, write_props,
};
use super::model::{Item, MAX_ITEMS, NewItem, NewProps, PROPS_CONTENT_TYPE, Props, Relationship};
use super::snapshot::{PartState, Snapshot};
use super::validation;

/// Discover and validate every explicit Custom XML Data Storage relationship.
pub fn discover(package: &OpcPackage) -> Result<Vec<Item>> {
    let mut items = Vec::new();
    scan(
        package,
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

/// Atomically add a validated data part, optional properties part, and relationships.
///
/// All fallible graph and XML work happens before package mutation. Defensive
/// rollback also covers an unexpected insertion or relationship failure, so an
/// error never exposes a partially-created Custom XML item.
pub fn add(package: &mut OpcPackage, item: NewItem) -> Result<()> {
    validate_content_type(&item.content_type)?;
    validate_payload(&item.xml)?;
    require_rel_id(&item.rel_id, "custom XML relationship")?;
    if item.part.as_str() == "/" {
        return invalid("custom XML data part cannot be the package root");
    }

    let source = package.get_part(&item.source).map_err(|error| {
        Error::Missing(format!(
            "custom XML source '{}': {error}",
            item.source.as_str()
        ))
    })?;
    if source.rels().get(&item.rel_id).is_some() {
        return invalid(format!(
            "relationship '{}' already exists on '{}'",
            item.rel_id,
            item.source.as_str()
        ));
    }
    package.validate_new_part_name(&item.part)?;

    if let Some(new_props) = &item.props {
        validate_props(&new_props.value)?;
        require_rel_id(&new_props.rel_id, "custom XML properties relationship")?;
        if new_props.part.as_str() == "/" {
            return invalid("custom XML properties part cannot be the package root");
        }
        package.validate_new_part_name(&new_props.part)?;
    }
    validate_new_names(&item.part, item.props.as_ref().map(|props| &props.part))?;

    if let Some(candidate_id) = item.props.as_ref().map(|props| props.value.id.as_str()) {
        scan(package, |_, _, _, _, _, _, props, _, _| {
            if props
                .as_ref()
                .is_some_and(|existing| candidate_id.eq_ignore_ascii_case(&existing.id))
            {
                return invalid(format!("duplicate custom XML itemID '{candidate_id}'"));
            }
            Ok(())
        })?;
    }

    let NewItem {
        source,
        rel_id,
        part,
        content_type,
        xml,
        props,
        conformance,
    } = item;
    let prepared_props = if let Some(NewProps {
        part,
        rel_id,
        value,
    }) = props
    {
        let xml = write_props(&value, conformance)?;
        Some((part, rel_id, xml))
    } else {
        None
    };

    let mut data = XmlPart::new(part.clone(), content_type, xml);
    if let Some((props_part, props_rel_id, _)) = prepared_props.as_ref() {
        let target = props_part.relative_ref(part.base_uri());
        data.rels_mut().try_add_relationship(
            conformance.props_relationship().into(),
            target,
            props_rel_id.clone(),
            TargetMode::Internal,
        )?;
    }

    let inserted_props = if let Some((props_part, _, props_xml)) = prepared_props {
        package.try_add_part(Box::new(XmlPart::new(
            props_part.clone(),
            PROPS_CONTENT_TYPE.into(),
            props_xml,
        )))?;
        Some(props_part)
    } else {
        None
    };

    if let Err(error) = package.try_add_part(Box::new(data)) {
        rollback_parts(package, &part, inserted_props.as_ref());
        return Err(error.into());
    }

    let target = part.relative_ref(source.base_uri());
    let relation_result =
        package
            .get_part_mut(&source)
            .map_err(Error::from)
            .and_then(|source_part| {
                source_part
                    .rels_mut()
                    .try_add_relationship(
                        conformance.relationship().into(),
                        target,
                        rel_id,
                        TargetMode::Internal,
                    )
                    .map(|_| ())
                    .map_err(Error::from)
            });
    if let Err(error) = relation_result {
        rollback_parts(package, &part, inserted_props.as_ref());
        return Err(error);
    }
    package.unsign();
    Ok(())
}

fn scan(
    package: &OpcPackage,
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
                .ok_or_else(|| limit("custom XML items", MAX_ITEMS, usize::MAX))?;
            require_at_most("custom XML items", occurrences, MAX_ITEMS)?;
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
            let data = package.get_part(&requested_name).map_err(|error| {
                Error::Missing(format!(
                    "custom XML part '{}': {error}",
                    requested_name.as_str()
                ))
            })?;
            let part = data.partname().clone();
            validate_content_type(data.content_type())?;
            let root = if let Some(root) = cached_roots.get(&part) {
                root.clone()
            } else {
                let root = validate_payload(data.blob())?;
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
    let part = package.get_part(&requested_name).map_err(|error| {
        Error::Missing(format!(
            "custom XML properties part '{}': {error}",
            requested_name.as_str()
        ))
    })?;
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
        let props = read_props(part.blob())?;
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

fn validate_new_names(part: &PackURI, props_part: Option<&PackURI>) -> Result<()> {
    let mut candidates = OpcPackage::new();
    candidates.try_add_part(Box::new(XmlPart::new(
        part.clone(),
        "application/xml".into(),
        Vec::new(),
    )))?;
    if let Some(props_part) = props_part {
        candidates.try_add_part(Box::new(XmlPart::new(
            props_part.clone(),
            PROPS_CONTENT_TYPE.into(),
            Vec::new(),
        )))?;
    }
    Ok(())
}

fn rollback_parts(package: &mut OpcPackage, part: &PackURI, props_part: Option<&PackURI>) {
    package.remove_part(part);
    if let Some(props_part) = props_part {
        package.remove_part(props_part);
    }
}

/// Publish a complete desired Custom XML occurrence set onto an already
/// source-checked package clone.
pub(crate) fn apply_items(
    package: &mut OpcPackage,
    before: &Snapshot,
    desired: &[Item],
) -> Result<()> {
    validation::items(desired)?;
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
        let current = package.get_part(source_name)?;
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
                return invalid(format!(
                    "custom XML relationship '{}' collides on '{}'",
                    item.rel_id(),
                    source_name.as_str()
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
            let mut state = before
                .source()
                .part(props_part)
                .cloned()
                .or_else(|| package.get_part(props_part).ok().map(PartState::capture))
                .unwrap_or_else(|| PartState {
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
        .collect::<Result<Vec<_>>>()?;
    let staged_names = staged
        .iter()
        .map(|part| part.name.clone())
        .collect::<Vec<_>>();

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
        if !part_is_referenced(package, &name) {
            package.remove_part(&name);
        }
    }
    package.unsign();
    Ok(())
}

fn reject_part_aliases(items: &[Item]) -> Result<()> {
    for item in items {
        if item.source() == item.part() {
            return invalid(format!(
                "custom XML source '{}' cannot also be its data part",
                item.source().as_str()
            ));
        }
        if let Some(props_part) = item.props_part()
            && (props_part == item.part() || props_part == item.source())
        {
            return invalid("custom XML properties part aliases an owning part");
        }
    }
    Ok(())
}

fn staged_state(staged: &mut Vec<PartState>, state: PartState) -> Result<()> {
    if let Some(existing) = staged.iter_mut().find(|part| part.name == state.name) {
        if *existing != state {
            return invalid(format!(
                "custom XML graph requires conflicting states for '{}'",
                state.name.as_str()
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
        if let Ok(existing) = package.get_part(&part.name) {
            if !belongs_to_source && !is_host {
                return invalid(format!(
                    "custom XML destination part '{}' is occupied",
                    part.name.as_str()
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
        } else if !belongs_to_source {
            package.validate_new_part_name(&part.name)?;
        } else {
            return Err(Error::Missing(format!(
                "custom XML source part '{}' disappeared",
                part.name.as_str()
            )));
        }
    }
    Ok(())
}

fn part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|part| &part == target)
    }) || package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|part| &part == target)
        })
    })
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
