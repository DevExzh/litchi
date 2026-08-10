//! Cross-field and cross-part validation for rich-value owners.

use super::model::{
    ArrayData, ArrayValueType, Bag, BagType, Bags, DxfComplement, Property, PropertyValue,
    RichValueData, RichValueRels, Structures, ValueType, XfComplement,
};
use super::{
    MAX_ARRAY_COLUMNS, MAX_ARRAY_ROWS, MAX_BAGS, MAX_ITEMS, MAX_RELATIONSHIPS, OFFICE_MAX_COUNT,
    bounded, bounded_nonempty, invalid, limit,
};
use crate::error::Result;
use caseless::Caseless;
use std::collections::HashSet;
use unicode_normalization::UnicodeNormalization;

/// Validate rich-value data, optionally against its structure and array parts.
pub fn validate_data(
    value: &RichValueData,
    structures: Option<&Structures>,
    arrays: Option<&ArrayData>,
) -> Result<()> {
    if value.values.len() > MAX_ITEMS {
        return Err(limit("rich values"));
    }
    if value.values.len() > OFFICE_MAX_COUNT as usize {
        return Err(limit("rich-value count"));
    }
    for (index, item) in value.values.iter().enumerate() {
        if let Some(structures) = structures {
            let structure = structures
                .values
                .get(item.structure as usize)
                .ok_or_else(|| invalid("rich-value structure index is out of range"))?;
            if item.values.len() != structure.keys.len() {
                return Err(invalid(
                    "rich-value value count does not match its structure",
                ));
            }
            for (key, raw) in structure.keys.iter().zip(&item.values) {
                bounded(raw, "rich-value value")?;
                match &key.value_type {
                    ValueType::RichValue => {
                        let reference = index_value(raw, "rich-value reference")?;
                        if reference >= index {
                            return Err(invalid(
                                "rich-value references must point to an earlier rich value",
                            ));
                        }
                    },
                    ValueType::Array => {
                        if let Some(arrays) = arrays {
                            let reference = index_value(raw, "rich-array reference")?;
                            if reference >= arrays.values.len() {
                                return Err(invalid("rich-array reference is out of range"));
                            }
                        }
                    },
                    ValueType::Number
                    | ValueType::Integer
                    | ValueType::Boolean
                    | ValueType::Error
                    | ValueType::Text
                    | ValueType::SupportingBag
                    | ValueType::Unknown(_) => {},
                }
            }
        } else {
            for raw in &item.values {
                bounded(raw, "rich-value value")?;
            }
        }
        if let Some(fallback) = &item.fallback {
            bounded(&fallback.value, "rich-value fallback")?;
        }
        validate_opaque(&item.opaque)?;
    }
    validate_opaque(&value.opaque)?;
    if let Some(extension) = &value.extension_list {
        validate_opaque_one(extension)?;
    }
    Ok(())
}

/// Validate rich-value structure definitions and case-insensitive key identity.
pub fn validate_structures(value: &Structures) -> Result<()> {
    if value.values.len() > MAX_ITEMS {
        return Err(limit("rich-value structures"));
    }
    if value.values.len() > OFFICE_MAX_COUNT as usize {
        return Err(limit("rich-value structure count"));
    }
    for structure in &value.values {
        bounded_nonempty(&structure.type_name, "rich-value structure type")?;
        if structure.keys.is_empty() || structure.keys.len() > MAX_ITEMS {
            return Err(invalid("rich-value structure requires at least one key"));
        }
        let mut names = HashSet::new();
        for key in &structure.keys {
            bounded_nonempty(&key.name, "rich-value key")?;
            if key.name.chars().count() > 255 {
                return Err(invalid("rich-value key exceeds 255 characters"));
            }
            if !names.insert(identity(&key.name)) {
                return Err(invalid(
                    "rich-value structure contains duplicate case-insensitive keys",
                ));
            }
        }
        validate_opaque(&structure.opaque)?;
    }
    validate_opaque(&value.opaque)?;
    if let Some(extension) = &value.extension_list {
        validate_opaque_one(extension)?;
    }
    Ok(())
}

/// Validate rich arrays and their row-major dimensions.
pub fn validate_arrays(value: &ArrayData) -> Result<()> {
    if value.values.len() > MAX_ITEMS || value.values.len() > OFFICE_MAX_COUNT as usize {
        return Err(limit("rich arrays"));
    }
    for array in &value.values {
        if !(1..=MAX_ARRAY_ROWS).contains(&array.rows) {
            return Err(invalid("rich-array row count is outside the Office range"));
        }
        if !(1..=MAX_ARRAY_COLUMNS).contains(&array.columns) {
            return Err(invalid(
                "rich-array column count is outside the Office range",
            ));
        }
        let expected = (array.rows as usize)
            .checked_mul(array.columns as usize)
            .ok_or_else(|| limit("rich-array values"))?;
        if array.values.len() != expected {
            return Err(invalid(
                "rich-array value count does not match its dimensions",
            ));
        }
        for item in &array.values {
            bounded(&item.value, "rich-array value")?;
            if matches!(
                &item.value_type,
                ArrayValueType::RichValue | ArrayValueType::Array
            ) {
                index_value(&item.value, "rich-array reference")?;
            }
        }
        validate_opaque(&array.opaque)?;
    }
    validate_opaque(&value.opaque)?;
    if let Some(extension) = &value.extension_list {
        validate_opaque_one(extension)?;
    }
    Ok(())
}

/// Validate feature-property-bag values, references, and the typed checkbox
/// and XF-control chains defined by MS-XLSX section 2.3.9.
pub fn validate_bags(value: &Bags) -> Result<()> {
    if value.values.len() > MAX_BAGS || value.values.len() > OFFICE_MAX_COUNT as usize {
        return Err(limit("feature property bags"));
    }
    if value
        .count
        .is_some_and(|count| count as usize != value.values.len())
    {
        return Err(invalid(
            "feature property bag count does not match its bags",
        ));
    }
    for extension in &value.bag_extensions {
        validate_opaque_one(extension)?;
    }
    for (index, bag) in value.values.iter().enumerate() {
        validate_bag(value, index, bag)?;
    }
    validate_opaque(&value.opaque)?;
    if let Some(extension) = &value.extension_list {
        validate_opaque_one(extension)?;
    }
    Ok(())
}

/// Validate a rich-value relationship part independently from its package.
pub fn validate_rich_value_rels(value: &RichValueRels) -> Result<()> {
    if value.ids.len() > MAX_RELATIONSHIPS {
        return Err(limit("rich-value relationships"));
    }
    let mut ids = HashSet::new();
    for id in &value.ids {
        relationship_id(id)?;
        if !ids.insert(id.as_str()) {
            return Err(invalid("rich-value relationships contain duplicate IDs"));
        }
    }
    validate_opaque(&value.opaque)?;
    if let Some(extension) = &value.extension_list {
        validate_opaque_one(extension)?;
    }
    Ok(())
}

/// Validate an xfComplement reference against the feature-bag graph.
pub fn validate_xf_complement(value: &XfComplement, bags: &Bags) -> Result<()> {
    if value.opaque.len() > MAX_ITEMS {
        return Err(limit("xfComplement extensions"));
    }
    validate_bags(bags)?;
    let mapping = mapping_index(bags, &BagType::XfComplements)?;
    if value.index as usize >= mapping.len() {
        return Err(invalid(
            "xfComplement index is outside MappedFeaturePropertyBags",
        ));
    }
    let bag_index = mapping[value.index as usize];
    if !matches!(&bags.values[bag_index].bag_type, BagType::XfComplement) {
        return Err(invalid(
            "xfComplement index does not identify an XFComplement bag",
        ));
    }
    validate_opaque(&value.opaque)
}

/// Validate a `DXFComplement` reference against the feature-bag graph.
pub fn validate_dxf_complement(value: &DxfComplement, bags: &Bags) -> Result<()> {
    if value.opaque.len() > MAX_ITEMS {
        return Err(limit("DXFComplement extensions"));
    }
    validate_bags(bags)?;
    let mapping = mapping_index(bags, &BagType::DxfComplements)?;
    if value.index as usize >= mapping.len() {
        return Err(invalid(
            "DXFComplement index is outside MappedFeaturePropertyBags",
        ));
    }
    let bag_index = mapping[value.index as usize];
    if !matches!(&bags.values[bag_index].bag_type, BagType::XfComplement) {
        return Err(invalid(
            "DXFComplement index does not identify an XFComplement bag",
        ));
    }
    validate_opaque(&value.opaque)
}

fn validate_bag(owner: &Bags, index: usize, bag: &Bag) -> Result<()> {
    bounded_nonempty(bag.bag_type.token(), "feature property bag type")?;
    if let Some(value) = &bag.ext_ref {
        bounded(value, "feature property bag extRef")?;
    }
    if let Some(value) = &bag.attribute {
        bounded(value, "feature property bag att")?;
    }
    if let Some(extension) = bag.bag_extension
        && extension as usize >= owner.bag_extensions.len()
    {
        return Err(invalid(
            "feature property bag extension index is out of range",
        ));
    }
    if bag.properties.len() > MAX_ITEMS {
        return Err(limit("feature property bag properties"));
    }
    let mut keys = HashSet::new();
    for property in &bag.properties {
        if let Some(key) = property.key() {
            bounded_nonempty(key, "feature property key")?;
            if key.chars().count() > 255 {
                return Err(invalid("feature property key exceeds 255 characters"));
            }
            if !keys.insert(identity(key)) {
                return Err(invalid(
                    "feature property bag contains duplicate case-insensitive keys",
                ));
            }
        }
        validate_property(owner, index, property)?;
    }
    validate_semantics(owner, bag)
}

fn validate_property(owner: &Bags, index: usize, property: &Property) -> Result<()> {
    match property {
        Property::Array { values, .. } => {
            if values.len() > MAX_ITEMS {
                return Err(limit("feature property array"));
            }
            for value in values {
                validate_property_value(owner, value)?;
            }
        },
        Property::Bag { index: target, .. } => {
            if *target as usize >= index || *target as usize >= owner.values.len() {
                return Err(invalid(
                    "feature property bag references must point to an earlier bag",
                ));
            }
        },
        Property::Integer { value, .. } => {
            integer(value)?;
        },
        Property::Text { value, .. } => bounded(value, "feature property text")?,
        Property::Boolean { .. } => {},
        Property::Decimal { value, .. } => double(value)?,
        Property::Relationship { id, .. } => relationship_id(id)?,
        Property::Unknown(value) => validate_opaque_one(value)?,
    }
    Ok(())
}

fn validate_property_value(owner: &Bags, value: &PropertyValue) -> Result<()> {
    match value {
        PropertyValue::Bag(index) => {
            if *index as usize >= owner.values.len() {
                return Err(invalid("feature property array bag index is out of range"));
            }
        },
        PropertyValue::Integer(value) => {
            integer(value)?;
        },
        PropertyValue::Text(value) => bounded(value, "feature property array text")?,
        PropertyValue::Boolean(_) => {},
        PropertyValue::Decimal(value) => double(value)?,
        PropertyValue::Relationship(id) => relationship_id(id)?,
        PropertyValue::Unknown(value) => validate_opaque_one(value)?,
    }
    Ok(())
}

fn validate_semantics(owner: &Bags, bag: &Bag) -> Result<()> {
    match &bag.bag_type {
        BagType::XfComplements => {
            validate_mapping(owner, bag, &BagType::XfComplement)?;
        },
        BagType::DxfComplements => {
            validate_mapping(owner, bag, &BagType::XfComplement)?;
        },
        BagType::XfComplement => {
            let target = required_bag(owner, bag, "XFControls")?;
            if !matches!(&owner.values[target].bag_type, BagType::XfControls) {
                return Err(invalid("XFControls must reference an XFControls bag"));
            }
        },
        BagType::XfControls => {
            let target = required_bag(owner, bag, "CellControl")?;
            if !matches!(&owner.values[target].bag_type, BagType::Checkbox) {
                return Err(invalid("CellControl must reference a Checkbox bag"));
            }
        },
        BagType::Checkbox => {
            if let Some(property) = bag.property("default") {
                let Property::Integer { value, .. } = property else {
                    return Err(invalid("Checkbox default must be an integer property"));
                };
                if !matches!(value.trim(), "0" | "1" | "2") {
                    return Err(invalid("Checkbox default must be 0, 1, or 2"));
                }
            }
        },
        BagType::Unknown(_) => {},
    }
    Ok(())
}

fn validate_mapping(owner: &Bags, bag: &Bag, expected: &BagType) -> Result<()> {
    let Some(Property::Array { values, .. }) = bag.property("MappedFeaturePropertyBags") else {
        return Err(invalid(
            "feature complement bag requires MappedFeaturePropertyBags",
        ));
    };
    if values.is_empty() {
        return Err(invalid(
            "MappedFeaturePropertyBags must contain at least one bag ID",
        ));
    }
    for value in values {
        let PropertyValue::Bag(index) = value else {
            return Err(invalid(
                "MappedFeaturePropertyBags must contain only bag IDs",
            ));
        };
        let Some(target) = owner.values.get(*index as usize) else {
            return Err(invalid("MappedFeaturePropertyBags bag ID is out of range"));
        };
        if !same_bag_type(&target.bag_type, expected) {
            return Err(invalid(
                "MappedFeaturePropertyBags contains an unexpected bag type",
            ));
        }
    }
    Ok(())
}

fn required_bag(owner: &Bags, bag: &Bag, key: &str) -> Result<usize> {
    let Some(Property::Bag { index, .. }) = bag.property(key) else {
        return Err(invalid(format!("feature property bag requires '{key}'")));
    };
    let target = *index as usize;
    if target >= owner.values.len() {
        return Err(invalid("feature property bag reference is out of range"));
    }
    Ok(target)
}

fn mapping_index(owner: &Bags, expected: &BagType) -> Result<Vec<usize>> {
    let bag = owner
        .values
        .iter()
        .find(|bag| same_bag_type(&bag.bag_type, expected))
        .ok_or_else(|| invalid("feature complement mapping bag is missing"))?;
    let Some(Property::Array { values, .. }) = bag.property("MappedFeaturePropertyBags") else {
        return Err(invalid(
            "feature complement mapping bag is missing MappedFeaturePropertyBags",
        ));
    };
    values
        .iter()
        .map(|value| match value {
            PropertyValue::Bag(index) => Ok(*index as usize),
            PropertyValue::Integer(_)
            | PropertyValue::Text(_)
            | PropertyValue::Boolean(_)
            | PropertyValue::Decimal(_)
            | PropertyValue::Relationship(_)
            | PropertyValue::Unknown(_) => Err(invalid(
                "feature complement mapping contains a non-bag value",
            )),
        })
        .collect()
}

fn same_bag_type(left: &BagType, right: &BagType) -> bool {
    left == right
}

fn identity(value: &str) -> String {
    value.chars().nfd().default_case_fold().nfd().collect()
}

fn index_value(value: &str, name: &str) -> Result<usize> {
    value
        .trim()
        .parse::<usize>()
        .map_err(|_source| invalid(format!("{name} must be a non-negative integer")))
}

fn integer(value: &str) -> Result<()> {
    let value = value.trim();
    let digits = value
        .strip_prefix('+')
        .or_else(|| value.strip_prefix('-'))
        .unwrap_or(value);
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(invalid("feature property integer has invalid lexical form"));
    }
    Ok(())
}

fn double(value: &str) -> Result<()> {
    let value = value.trim();
    if matches!(value, "INF" | "-INF" | "NaN") || value.parse::<f64>().is_ok() {
        Ok(())
    } else {
        Err(invalid("feature property decimal has invalid lexical form"))
    }
}

fn relationship_id(value: &str) -> Result<()> {
    bounded_nonempty(value, "relationship ID")?;
    if value.chars().any(char::is_whitespace) {
        return Err(invalid("relationship ID cannot contain whitespace"));
    }
    Ok(())
}

fn validate_opaque(values: &[super::model::Opaque]) -> Result<()> {
    for value in values {
        validate_opaque_one(value)?;
    }
    Ok(())
}

fn validate_opaque_one(value: &super::model::Opaque) -> Result<()> {
    if value.xml.len() > super::MAX_OPAQUE_BYTES {
        return Err(limit("opaque XML"));
    }
    Ok(())
}
