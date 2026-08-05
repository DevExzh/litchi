//! OPC package graph discovery and mutation for Custom XML Data Storage.

use crate::mce::Name;
use crate::{Error, Result};
use litchi_opc::part::XmlPart;
use litchi_opc::{OpcPackage, PackURI, Part, TargetMode};
use std::collections::HashMap;

use super::codec::{
    invalid, is_data_relationship, is_props_relationship, limit, read_props, require_at_most,
    require_rel_id, validate_content_type, validate_payload, validate_props, write_props,
};
use super::model::{Item, MAX_ITEMS, NewItem, NewProps, PROPS_CONTENT_TYPE, Props};

/// Discover and validate every explicit Custom XML Data Storage relationship.
pub fn discover(package: &OpcPackage) -> Result<Vec<Item>> {
    let mut items = Vec::new();
    scan(
        package,
        |source, rel_id, part, data, root, props_part, props| {
            items.push(Item::new(
                source.clone(),
                rel_id.into(),
                part,
                data.content_type().into(),
                root,
                data.blob_arc(),
                props_part,
                props,
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
        scan(package, |_, _, _, _, _, _, props| {
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
        &str,
        PackURI,
        &dyn Part,
        Name,
        Option<PackURI>,
        Option<Props>,
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
            let (props_part, props) = resolve_props(
                package,
                data,
                &mut property_ids,
                &mut cached_props,
                &mut props_owners,
            )?;
            visit(
                source.partname(),
                relationship.r_id(),
                part,
                data,
                root,
                props_part,
                props,
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
) -> Result<(Option<PackURI>, Option<Props>)> {
    let mut relationships = data.rels().iter();
    let Some(relationship) = relationships.next() else {
        return Ok((None, None));
    };
    if relationships.next().is_some() {
        return invalid(format!(
            "custom XML part '{}' has more than one outbound relationship",
            data.partname().as_str()
        ));
    }
    if !is_props_relationship(relationship.reltype()) {
        return Err(Error::Relationship(format!(
            "custom XML part '{}' has forbidden relationship type '{}'",
            data.partname().as_str(),
            relationship.reltype()
        )));
    }
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
        if !part.rels().is_empty() {
            return invalid(format!(
                "custom XML properties part '{}' must not have relationships",
                part_name.as_str()
            ));
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
    Ok((Some(part_name), Some(props)))
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
