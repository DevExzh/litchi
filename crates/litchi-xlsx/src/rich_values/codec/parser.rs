//! Namespace-aware parsers for the supported rich-value XML parts.

use super::xml::{
    Node, no_attributes, opaque, optional, parse_document, require, required, whitespace,
};
use crate::error::Result;
use crate::rich_values::model::{
    Array, ArrayData, ArrayValue, ArrayValueType, Bag, BagType, Bags, DxfComplement, Fallback,
    FallbackType, Key, Opaque, Property, PropertyValue, RichValue, RichValueData, RichValueRels,
    Structure, Structures, ValueType, XfComplement,
};
use crate::rich_values::package::{Document, Kind};
use crate::rich_values::validation::{
    validate_arrays, validate_bags, validate_data, validate_rich_value_rels, validate_structures,
};
use crate::rich_values::{
    FEATURE_BAG, RELATIONSHIPS, RICH_DATA, RICH_DATA_2, RICH_VALUE_REL, SPREADSHEETML, invalid,
    limit,
};

const MAX_CHILDREN: usize = 1_000_000;

/// Parse an rvData part.
pub fn parse_data(xml: &[u8]) -> Result<RichValueData> {
    let root = parse_document(xml)?;
    require(&root, RICH_DATA, "rvData")?;
    no_attributes(&root, &[("", "count")])?;
    let count = required_u32(&root, "count")?;
    let mut values = Vec::new();
    let mut extension_list = None;
    let mut opaque_values = Vec::new();
    let mut extension_seen = false;
    for child in &root.children {
        match (child.namespace.as_str(), child.name.as_str()) {
            (RICH_DATA, "rv") if !extension_seen => {
                push_limit(&mut values, MAX_CHILDREN, "rich values")?;
                values.push(parse_value(child)?);
            },
            (SPREADSHEETML, "extLst") if !extension_seen => {
                extension_seen = true;
                extension_list = Some(opaque(child)?);
            },
            _ => opaque_values.push(opaque(child)?),
        }
    }
    if count as usize != values.len() {
        return Err(invalid("rvData count does not match its rich values"));
    }
    let value = RichValueData {
        values,
        extension_list,
        opaque: opaque_values,
    };
    validate_data(&value, None, None)?;
    Ok(value)
}

/// Parse an rvStructures part.
pub fn parse_structures(xml: &[u8]) -> Result<Structures> {
    let root = parse_document(xml)?;
    require(&root, RICH_DATA, "rvStructures")?;
    no_attributes(&root, &[("", "count")])?;
    let count = required_u32(&root, "count")?;
    let mut values = Vec::new();
    let mut extension_list = None;
    let mut opaque_values = Vec::new();
    let mut extension_seen = false;
    for child in &root.children {
        match (child.namespace.as_str(), child.name.as_str()) {
            (RICH_DATA, "s") if !extension_seen => {
                push_limit(&mut values, MAX_CHILDREN, "rich-value structures")?;
                values.push(parse_structure(child)?);
            },
            (SPREADSHEETML, "extLst") if !extension_seen => {
                extension_seen = true;
                extension_list = Some(opaque(child)?);
            },
            _ => opaque_values.push(opaque(child)?),
        }
    }
    if count as usize != values.len() {
        return Err(invalid("rvStructures count does not match its structures"));
    }
    let value = Structures {
        values,
        extension_list,
        opaque: opaque_values,
    };
    validate_structures(&value)?;
    Ok(value)
}

/// Parse an arrayData part.
pub fn parse_arrays(xml: &[u8]) -> Result<ArrayData> {
    let root = parse_document(xml)?;
    require(&root, RICH_DATA_2, "arrayData")?;
    no_attributes(&root, &[("", "count")])?;
    let count = required_u32(&root, "count")?;
    let mut values = Vec::new();
    let mut extension_list = None;
    let mut opaque_values = Vec::new();
    let mut extension_seen = false;
    for child in &root.children {
        match (child.namespace.as_str(), child.name.as_str()) {
            (RICH_DATA_2, "a") if !extension_seen => {
                push_limit(&mut values, MAX_CHILDREN, "rich arrays")?;
                values.push(parse_array(child)?);
            },
            (SPREADSHEETML, "extLst") if !extension_seen => {
                extension_seen = true;
                extension_list = Some(opaque(child)?);
            },
            _ => opaque_values.push(opaque(child)?),
        }
    }
    if count as usize != values.len() {
        return Err(invalid("arrayData count does not match its arrays"));
    }
    let value = ArrayData {
        values,
        extension_list,
        opaque: opaque_values,
    };
    validate_arrays(&value)?;
    Ok(value)
}

/// Parse a `FeaturePropertyBags` part.
pub fn parse_feature_property_bags(xml: &[u8]) -> Result<Bags> {
    let root = parse_document(xml)?;
    require(&root, FEATURE_BAG, "FeaturePropertyBags")?;
    no_attributes(&root, &[("", "count")])?;
    let count = optional(&root, "", "count")
        .map(|value| parse_u32(value, "bag count"))
        .transpose()?;
    let mut bag_extensions = Vec::new();
    let mut values = Vec::new();
    let mut extension_list = None;
    let mut opaque_values = Vec::new();
    let mut stage = 0u8;
    for child in &root.children {
        match (child.namespace.as_str(), child.name.as_str()) {
            (FEATURE_BAG, "bagExt") if stage == 0 => {
                push_limit(
                    &mut bag_extensions,
                    MAX_CHILDREN,
                    "feature property bag extensions",
                )?;
                bag_extensions.push(opaque(child)?);
            },
            (FEATURE_BAG, "bag") if stage <= 1 => {
                stage = 1;
                push_limit(&mut values, MAX_CHILDREN, "feature property bags")?;
                values.push(parse_bag(child)?);
            },
            (SPREADSHEETML, "extLst") if stage <= 2 => {
                stage = 2;
                extension_list = Some(opaque(child)?);
            },
            _ => opaque_values.push(opaque(child)?),
        }
    }
    let value = Bags {
        count,
        bag_extensions,
        values,
        extension_list,
        opaque: opaque_values,
    };
    validate_bags(&value)?;
    Ok(value)
}

/// Parse a richValueRels part.
pub fn parse_rich_value_rels(xml: &[u8]) -> Result<RichValueRels> {
    let root = parse_document(xml)?;
    require(&root, RICH_VALUE_REL, "richValueRels")?;
    no_attributes(&root, &[])?;
    let mut ids = Vec::new();
    let mut extension_list = None;
    let mut opaque_values = Vec::new();
    let mut extension_seen = false;
    for child in &root.children {
        match (child.namespace.as_str(), child.name.as_str()) {
            (RICH_VALUE_REL, "rel") if !extension_seen => {
                no_attributes(
                    child,
                    &[
                        (RELATIONSHIPS, "id"),
                        (crate::rich_values::STRICT_RELATIONSHIPS, "id"),
                    ],
                )?;
                if !child.children.is_empty() || !child.text.is_empty() {
                    return Err(invalid("richValueRels rel must be empty"));
                }
                ids.push(
                    optional(child, RELATIONSHIPS, "id")
                        .or_else(|| optional(child, crate::rich_values::STRICT_RELATIONSHIPS, "id"))
                        .ok_or_else(|| invalid("richValueRels rel is missing r:id"))?
                        .to_owned(),
                );
            },
            (SPREADSHEETML, "extLst") if !extension_seen => {
                extension_seen = true;
                extension_list = Some(opaque(child)?);
            },
            _ => opaque_values.push(opaque(child)?),
        }
    }
    let value = RichValueRels {
        ids,
        extension_list,
        opaque: opaque_values,
    };
    validate_rich_value_rels(&value)?;
    Ok(value)
}

/// Parse one xfComplement extension element.
pub fn parse_xf_complement(xml: &[u8]) -> Result<XfComplement> {
    let root = parse_document(xml)?;
    require(&root, FEATURE_BAG, "xfComplement")?;
    no_attributes(&root, &[("", "i")])?;
    whitespace(&root)?;
    let value = XfComplement {
        index: required_u32(&root, "i")?,
        opaque: root
            .children
            .iter()
            .map(opaque)
            .collect::<Result<Vec<_>>>()?,
    };
    Ok(value)
}

/// Parse one `DXFComplement` extension element.
pub fn parse_dxf_complement(xml: &[u8]) -> Result<DxfComplement> {
    let root = parse_document(xml)?;
    require(&root, FEATURE_BAG, "DXFComplement")?;
    no_attributes(&root, &[("", "i")])?;
    whitespace(&root)?;
    let value = DxfComplement {
        index: required_u32(&root, "i")?,
        opaque: root
            .children
            .iter()
            .map(opaque)
            .collect::<Result<Vec<_>>>()?,
    };
    Ok(value)
}

pub(crate) fn parse_part(kind: Kind, xml: &[u8]) -> Result<Document> {
    Ok(match kind {
        Kind::Data => Document::Data(parse_data(xml)?),
        Kind::Structures => Document::Structures(parse_structures(xml)?),
        Kind::Arrays => Document::Arrays(parse_arrays(xml)?),
        Kind::Relationships => Document::Relationships(parse_rich_value_rels(xml)?),
        Kind::FeatureBags => Document::FeatureBags(parse_feature_property_bags(xml)?),
        Kind::Styles
        | Kind::SupportingData
        | Kind::SupportingStructures
        | Kind::Types
        | Kind::WebImages => Document::Opaque(Opaque::new(xml.to_vec())?),
    })
}

fn parse_value(node: &Node) -> Result<RichValue> {
    no_attributes(node, &[("", "s")])?;
    let structure = required_u32(node, "s")?;
    let mut fallback = None;
    let mut values = Vec::new();
    let mut opaque_values = Vec::new();
    let mut value_started = false;
    let mut fallback_seen = false;
    for child in &node.children {
        match (child.namespace.as_str(), child.name.as_str()) {
            (RICH_DATA, "fb") if !fallback_seen && !value_started => {
                fallback_seen = true;
                fallback = Some(parse_fallback(child)?);
            },
            (RICH_DATA, "v") => {
                value_started = true;
                no_attributes(child, &[])?;
                if !child.children.is_empty() {
                    return Err(invalid("rich-value v must be a leaf"));
                }
                values.push(child.text.clone());
            },
            _ => opaque_values.push(opaque(child)?),
        }
    }
    if values.is_empty() {
        return Err(invalid("rich-value requires at least one v element"));
    }
    Ok(RichValue {
        structure,
        fallback,
        values,
        opaque: opaque_values,
    })
}

fn parse_fallback(node: &Node) -> Result<Fallback> {
    no_attributes(node, &[("", "t")])?;
    if !node.children.is_empty() {
        return Err(invalid("rich-value fallback must be a leaf"));
    }
    let value_type = optional(node, "", "t")
        .map(FallbackType::parse)
        .transpose()?
        .unwrap_or(FallbackType::Number);
    Ok(Fallback {
        value_type,
        value: node.text.clone(),
    })
}

fn parse_structure(node: &Node) -> Result<Structure> {
    no_attributes(node, &[("", "t")])?;
    let type_name = required(node, "", "t")?.to_owned();
    let mut keys = Vec::new();
    let mut opaque_values = Vec::new();
    for child in &node.children {
        if child.namespace == RICH_DATA && child.name == "k" {
            keys.push(parse_key(child)?);
        } else {
            opaque_values.push(opaque(child)?);
        }
    }
    if keys.is_empty() {
        return Err(invalid("rich-value structure requires at least one key"));
    }
    Ok(Structure {
        type_name,
        keys,
        opaque: opaque_values,
    })
}

fn parse_key(node: &Node) -> Result<Key> {
    no_attributes(node, &[("", "n"), ("", "t")])?;
    if !node.children.is_empty() || !node.text.is_empty() {
        return Err(invalid("rich-value key must be empty"));
    }
    Ok(Key {
        name: required(node, "", "n")?.to_owned(),
        value_type: optional(node, "", "t").map_or(ValueType::Number, ValueType::parse),
    })
}

fn parse_array(node: &Node) -> Result<Array> {
    no_attributes(node, &[("", "r"), ("", "c")])?;
    let rows = required_u32(node, "r")?;
    let columns = optional(node, "", "c")
        .map(|value| parse_u32(value, "rich-array columns"))
        .transpose()?
        .unwrap_or(1);
    let mut values = Vec::new();
    let mut opaque_values = Vec::new();
    for child in &node.children {
        if child.namespace == RICH_DATA_2 && child.name == "v" {
            values.push(parse_array_value(child)?);
        } else {
            opaque_values.push(opaque(child)?);
        }
    }
    Ok(Array {
        rows,
        columns,
        values,
        opaque: opaque_values,
    })
}

fn parse_array_value(node: &Node) -> Result<ArrayValue> {
    no_attributes(node, &[("", "t")])?;
    if !node.children.is_empty() {
        return Err(invalid("rich-array value must be a leaf"));
    }
    Ok(ArrayValue {
        value_type: optional(node, "", "t").map_or(ArrayValueType::Number, ArrayValueType::parse),
        value: node.text.clone(),
    })
}

fn parse_bag(node: &Node) -> Result<Bag> {
    no_attributes(
        node,
        &[("", "type"), ("", "extRef"), ("", "bagExtId"), ("", "att")],
    )?;
    let bag_type = BagType::parse(required(node, "", "type")?);
    let bag_extension = optional(node, "", "bagExtId")
        .map(|value| parse_u32(value, "bagExtId"))
        .transpose()?;
    let mut properties = Vec::new();
    let mut opaque_values = Vec::new();
    for child in &node.children {
        if child.namespace == FEATURE_BAG
            && matches!(
                child.name.as_str(),
                "a" | "bagId" | "i" | "s" | "b" | "d" | "rel"
            )
        {
            properties.push(parse_property(child)?);
        } else {
            opaque_values.push(opaque(child)?);
        }
    }
    Ok(Bag {
        bag_type,
        ext_ref: optional(node, "", "extRef").map(str::to_owned),
        bag_extension,
        attribute: optional(node, "", "att").map(str::to_owned),
        properties,
        opaque: opaque_values,
    })
}

fn parse_property(node: &Node) -> Result<Property> {
    match node.name.as_str() {
        "a" => {
            no_attributes(node, &[("", "k")])?;
            let mut values = Vec::new();
            for child in &node.children {
                values.push(parse_property_value(child)?);
            }
            Ok(Property::Array {
                key: required(node, "", "k")?.to_owned(),
                values,
            })
        },
        "bagId" => Ok(Property::Bag {
            key: required(node, "", "k")?.to_owned(),
            index: leaf_u32(node, "k")?,
        }),
        "i" => Ok(Property::Integer {
            key: required(node, "", "k")?.to_owned(),
            value: leaf_text(node, "k")?,
        }),
        "s" => Ok(Property::Text {
            key: required(node, "", "k")?.to_owned(),
            value: leaf_text(node, "k")?,
        }),
        "b" => Ok(Property::Boolean {
            key: required(node, "", "k")?.to_owned(),
            value: leaf_bool(node, "k")?,
        }),
        "d" => Ok(Property::Decimal {
            key: required(node, "", "k")?.to_owned(),
            value: leaf_text(node, "k")?,
        }),
        "rel" => Ok(Property::Relationship {
            key: required(node, "", "k")?.to_owned(),
            id: leaf_text(node, "k")?,
        }),
        _ => Err(invalid("unsupported feature property element")),
    }
}

fn parse_property_value(node: &Node) -> Result<PropertyValue> {
    if node.namespace != FEATURE_BAG {
        return Ok(PropertyValue::Unknown(opaque(node)?));
    }
    if !node.children.is_empty() {
        return Ok(PropertyValue::Unknown(opaque(node)?));
    }
    Ok(match node.name.as_str() {
        "bagId" => PropertyValue::Bag(leaf_u32(node, "bagId")?),
        "i" => PropertyValue::Integer(leaf_text(node, "i")?),
        "s" => PropertyValue::Text(leaf_text(node, "s")?),
        "b" => PropertyValue::Boolean(leaf_bool(node, "b")?),
        "d" => PropertyValue::Decimal(leaf_text(node, "d")?),
        "rel" => PropertyValue::Relationship(leaf_text(node, "rel")?),
        _ => PropertyValue::Unknown(opaque(node)?),
    })
}

fn leaf_u32(node: &Node, key: &str) -> Result<u32> {
    no_attributes(node, &[("", key)])?;
    if !node.children.is_empty() {
        return Err(invalid("feature property scalar must be a leaf"));
    }
    parse_u32(&node.text, key)
}

fn leaf_text(node: &Node, key: &str) -> Result<String> {
    no_attributes(node, &[("", key)])?;
    if !node.children.is_empty() {
        return Err(invalid("feature property scalar must be a leaf"));
    }
    Ok(node.text.trim().to_owned())
}

fn leaf_bool(node: &Node, key: &str) -> Result<bool> {
    match leaf_text(node, key)?.as_str() {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid("feature property boolean has invalid lexical form")),
    }
}

fn required_u32(node: &Node, name: &str) -> Result<u32> {
    parse_u32(required(node, "", name)?, name)
}

fn parse_u32(value: &str, name: &str) -> Result<u32> {
    value
        .trim()
        .parse()
        .map_err(|_| invalid(format!("{name} must be an unsigned integer")))
}

fn push_limit<T>(values: &mut [T], limit_value: usize, name: &str) -> Result<()> {
    if values.len() >= limit_value {
        Err(limit(name))
    } else {
        Ok(())
    }
}
