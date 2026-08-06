//! OPC discovery and graph validation for rich-value parts.

use super::model::{Document, Kind, Package, Part};
use crate::error::Result;
use crate::rich_values::codec;
use crate::rich_values::model::{Link, Mode};
use crate::rich_values::validation::{
    validate_arrays, validate_bags, validate_data, validate_rich_value_rels, validate_structures,
};
use crate::rich_values::{MAX_RELATIONSHIPS, invalid, limit};
use litchi_opc::{OpcPackage, Relationships};

/// Discover, parse, and validate every rich-value family part in an OPC package.
///
/// Non-rich parts are not copied into this owner. Their relationship edges are,
/// however, included in [`Package::topology`] so incoming and outgoing graph
/// topology remains observable without following or executing any target.
pub fn load(package: &OpcPackage) -> Result<Package> {
    let package_relationships = collect_links(None, package.rels())?;
    let mut topology = package_relationships.clone();
    let mut parts = Vec::new();
    let mut seen = [false; 10];

    for source in package.iter_parts() {
        let Some(kind) = Kind::from_content_type(source.content_type()) else {
            topology.extend(collect_links(
                Some(source.partname().as_str()),
                source.rels(),
            )?);
            continue;
        };
        let slot = kind as usize;
        if seen[slot] {
            return Err(invalid(format!(
                "XLSX rich-values package contains multiple '{}' parts",
                kind.content_type()
            )));
        }
        seen[slot] = true;
        let relationships = collect_links(Some(source.partname().as_str()), source.rels())?;
        topology.extend(relationships.iter().cloned());
        parts.push(Part {
            name: source.partname().as_str().to_owned(),
            kind,
            document: codec::parse_part(kind, source.blob())?,
            relationships,
        });
    }

    parts.sort_by(|left, right| left.name.cmp(&right.name));
    topology.sort_by(|left, right| {
        left.source
            .cmp(&right.source)
            .then_with(|| left.id.cmp(&right.id))
    });
    validate_parts(&parts)?;
    validate_rich_relationship_ids(&parts)?;
    Ok(Package {
        relationships: package_relationships,
        topology,
        parts,
    })
}

fn collect_links(source: Option<&str>, relationships: &Relationships) -> Result<Vec<Link>> {
    let mut result = Vec::new();
    for relationship in relationships.iter() {
        if result.len() >= MAX_RELATIONSHIPS {
            return Err(limit("relationship edges"));
        }
        let (target, resolved_target, mode) = if relationship.is_external() {
            (relationship.target_ref().to_owned(), None, Mode::External)
        } else {
            let resolved = relationship.target_partname()?.as_str().to_owned();
            (
                relationship.target_ref().to_owned(),
                Some(resolved),
                Mode::Internal,
            )
        };
        result.push(Link {
            source: source.map(str::to_owned),
            id: relationship.r_id().to_owned(),
            relationship_type: relationship.reltype().to_owned(),
            target,
            resolved_target,
            mode,
        });
    }
    result.sort_by(|left, right| left.id.cmp(&right.id));
    Ok(result)
}

fn validate_parts(parts: &[Part]) -> Result<()> {
    let has = |kind| parts.iter().any(|part| part.kind == kind);
    if has(Kind::Data) && !has(Kind::Structures) {
        return Err(invalid("rich-value data requires a structure part"));
    }
    if has(Kind::Structures) && !has(Kind::Data) {
        return Err(invalid("rich-value structures require a data part"));
    }
    for required in [
        Kind::Arrays,
        Kind::Relationships,
        Kind::Types,
        Kind::WebImages,
    ] {
        if has(required) && (!has(Kind::Data) || !has(Kind::Structures)) {
            return Err(invalid(format!(
                "'{}' requires rich-value data and structures",
                required.content_type()
            )));
        }
    }
    if has(Kind::SupportingData) != has(Kind::SupportingStructures) {
        return Err(invalid(
            "supporting property-bag data and structures must be paired",
        ));
    }
    if has(Kind::SupportingData) && (!has(Kind::Data) || !has(Kind::Structures)) {
        return Err(invalid(
            "supporting property-bag parts require rich-value data and structures",
        ));
    }
    if has(Kind::Styles) && (!has(Kind::SupportingData) || !has(Kind::SupportingStructures)) {
        return Err(invalid(
            "rich styles require supporting property-bag data and structures",
        ));
    }

    let data = parts.iter().find_map(|part| match &part.document {
        Document::Data(value) => Some(value),
        _ => None,
    });
    let structures = parts.iter().find_map(|part| match &part.document {
        Document::Structures(value) => Some(value),
        _ => None,
    });
    let arrays = parts.iter().find_map(|part| match &part.document {
        Document::Arrays(value) => Some(value),
        _ => None,
    });
    if let Some(value) = data {
        validate_data(value, structures, arrays)?;
    }
    if let Some(value) = structures {
        validate_structures(value)?;
    }
    if let Some(value) = arrays {
        validate_arrays(value)?;
    }
    if let Some(value) = parts.iter().find_map(|part| match &part.document {
        Document::Relationships(value) => Some(value),
        _ => None,
    }) {
        validate_rich_value_rels(value)?;
    }
    if let Some(value) = parts.iter().find_map(|part| match &part.document {
        Document::FeatureBags(value) => Some(value),
        _ => None,
    }) {
        validate_bags(value)?;
    }
    Ok(())
}

fn validate_rich_relationship_ids(parts: &[Part]) -> Result<()> {
    let Some(part) = parts.iter().find(|part| part.kind == Kind::Relationships) else {
        return Ok(());
    };
    let Document::Relationships(value) = &part.document else {
        return Ok(());
    };
    for id in &value.ids {
        if !part.relationships.iter().any(|link| link.id == *id) {
            return Err(invalid(format!(
                "richValueRels references missing relationship '{id}'"
            )));
        }
    }
    Ok(())
}
