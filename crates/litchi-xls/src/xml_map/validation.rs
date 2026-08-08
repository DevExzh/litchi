//! Cross-object XML-map invariants and list-column dependencies.

use super::model::{MapInfo, SchemaId};
use crate::list_object::ListObjectSourceMetadata;
use crate::{Error, Result};
use std::collections::HashSet;

pub(crate) fn validate(value: &MapInfo) -> Result<()> {
    let mut schema_ids = HashSet::with_capacity(value.schemas().len());
    for schema in value.schemas() {
        if !schema_ids.insert(schema.id()) {
            return Err(invalid("duplicate Schema ID"));
        }
        let mut references = HashSet::new();
        if let Some(items) = schema.schema_references() {
            for reference in items {
                if !references.insert(reference) {
                    return Err(invalid("SchemaRef contains a duplicate schema ID"));
                }
                if reference == schema.id() {
                    return Err(invalid("SchemaRef cannot reference its own Schema ID"));
                }
                if !schema_ids.contains(reference) {
                    // The complete set is checked below; this early check is
                    // intentionally omitted for forward schema references.
                }
            }
        }
        validate_namespace_declarations(schema.namespaces())?;
    }
    for schema in value.schemas() {
        if let Some(items) = schema.schema_references() {
            for reference in items {
                if !schema_ids.contains(reference) {
                    return Err(invalid("SchemaRef points to a missing Schema ID"));
                }
            }
        }
    }
    validate_schema_cycles(value)?;

    let mut map_ids = HashSet::with_capacity(value.maps().len());
    let mut map_names = HashSet::with_capacity(value.maps().len());
    let mut binding_names = HashSet::new();
    let mut file_binding_names = HashSet::new();
    for map in value.maps() {
        if !map_ids.insert(map.id()) {
            return Err(invalid("duplicate Map ID"));
        }
        if !map_names.insert(map.name()) {
            return Err(invalid("duplicate Map Name"));
        }
        if !schema_ids.contains(map.schema_id()) {
            return Err(invalid("Map SchemaID points to a missing Schema ID"));
        }
        validate_namespace_declarations(map.namespaces())?;
        if let Some(binding) = map.data_binding() {
            validate_namespace_declarations(binding.namespaces())?;
            if let Some(name) = binding.data_binding_name() {
                if !binding_names.insert(name) {
                    return Err(invalid("duplicate DataBindingName"));
                }
            }
            if let Some(name) = binding.file_binding_name() {
                if !file_binding_names.insert(name) {
                    return Err(invalid("duplicate FileBindingName"));
                }
            }
        }
    }
    validate_namespace_declarations(value.namespaces())
}

pub(crate) fn validate_list_columns(
    value: Option<&MapInfo>,
    worksheets: &[crate::worksheet::Worksheet],
) -> Result<()> {
    validate_list_objects(
        value,
        worksheets
            .iter()
            .flat_map(crate::worksheet::Worksheet::list_objects),
    )
}

pub(crate) fn validate_list_objects<'a>(
    value: Option<&MapInfo>,
    tables: impl IntoIterator<Item = &'a crate::list_object::ListObject>,
) -> Result<()> {
    for table in tables {
        let Some(ListObjectSourceMetadata::Xml(metadata)) = table.source_metadata() else {
            continue;
        };
        for field in metadata.fields() {
            let Some(mapping) = field.mapping() else {
                continue;
            };
            let Some(value) = value else {
                return Err(invalid("XML column mapping has no XML MapInfo stream"));
            };
            if value.map(mapping.map_identifier()).is_none() {
                return Err(invalid(format!(
                    "XML column mapping references missing Map ID {}",
                    mapping.map_identifier()
                )));
            }
        }
    }
    Ok(())
}

fn validate_namespace_declarations(values: &[super::model::NamespaceDeclaration]) -> Result<()> {
    let mut prefixes = HashSet::new();
    for value in values {
        if !prefixes.insert(value.prefix()) {
            return Err(invalid("duplicate namespace declaration prefix"));
        }
    }
    Ok(())
}

fn validate_schema_cycles(value: &MapInfo) -> Result<()> {
    for schema in value.schemas() {
        let mut path = Vec::new();
        if reaches(schema.id(), schema.id(), value, &mut path)? {
            return Err(invalid("SchemaRef dependency graph contains a cycle"));
        }
    }
    Ok(())
}

fn reaches(
    origin: &SchemaId,
    current: &SchemaId,
    value: &MapInfo,
    path: &mut Vec<SchemaId>,
) -> Result<bool> {
    let Some(schema) = value.schema(current) else {
        return Err(invalid("SchemaRef points to a missing Schema ID"));
    };
    let Some(references) = schema.schema_references() else {
        return Ok(false);
    };
    for reference in references {
        if reference == origin {
            return Ok(true);
        }
        if path.iter().any(|item| item == reference) {
            return Ok(true);
        }
        path.push(reference.clone());
        let found = reaches(origin, reference, value, path)?;
        path.pop();
        if found {
            return Ok(true);
        }
    }
    Ok(false)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidData(format!("XML map: {}", message.into()))
}
