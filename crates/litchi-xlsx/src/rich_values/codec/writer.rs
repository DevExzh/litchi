//! Deterministic writers for the typed rich-value XML parts.

use super::super::model::{
    Array, ArrayData, ArrayValue, Bag, Bags, DxfComplement, Fallback, Property, PropertyValue,
    RichValue, RichValueData, RichValueRels, Structure, Structures, XfComplement,
};
use super::super::package::{Document, Kind};
use super::super::validation::{
    validate_arrays, validate_bags, validate_data, validate_rich_value_rels, validate_structures,
};
use super::super::{
    FEATURE_BAG, MAX_OUTPUT_BYTES, RELATIONSHIPS, RICH_DATA, RICH_DATA_2, RICH_VALUE_REL,
    SPREADSHEETML, invalid, limit,
};
use super::xml::{attr, escape_attr, escape_text};
use crate::error::Result;

/// Write an rvData part.
pub fn write_data(value: &RichValueData) -> Result<Vec<u8>> {
    validate_data(value, None, None)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<rd:rvData xmlns:rd=\"");
    escape_attr(&mut output, RICH_DATA);
    output.extend_from_slice(b"\" xmlns:x=\"");
    escape_attr(&mut output, SPREADSHEETML);
    output.push(b'"');
    attr(&mut output, "count", &value.values.len().to_string());
    output.push(b'>');
    for item in &value.values {
        write_value(&mut output, item)?;
    }
    if let Some(extension) = &value.extension_list {
        output.extend_from_slice(&extension.xml);
    }
    for item in &value.opaque {
        output.extend_from_slice(&item.xml);
    }
    output.extend_from_slice(b"</rd:rvData>");
    finish(output)
}

/// Write an rvStructures part.
pub fn write_structures(value: &Structures) -> Result<Vec<u8>> {
    validate_structures(value)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<rd:rvStructures xmlns:rd=\"");
    escape_attr(&mut output, RICH_DATA);
    output.extend_from_slice(b"\" xmlns:x=\"");
    escape_attr(&mut output, SPREADSHEETML);
    output.push(b'"');
    attr(&mut output, "count", &value.values.len().to_string());
    output.push(b'>');
    for item in &value.values {
        write_structure(&mut output, item)?;
    }
    if let Some(extension) = &value.extension_list {
        output.extend_from_slice(&extension.xml);
    }
    for item in &value.opaque {
        output.extend_from_slice(&item.xml);
    }
    output.extend_from_slice(b"</rd:rvStructures>");
    finish(output)
}

/// Write an arrayData part.
pub fn write_arrays(value: &ArrayData) -> Result<Vec<u8>> {
    validate_arrays(value)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<rd2:arrayData xmlns:rd2=\"");
    escape_attr(&mut output, RICH_DATA_2);
    output.extend_from_slice(b"\" xmlns:x=\"");
    escape_attr(&mut output, SPREADSHEETML);
    output.push(b'"');
    attr(&mut output, "count", &value.values.len().to_string());
    output.push(b'>');
    for item in &value.values {
        write_array(&mut output, item)?;
    }
    if let Some(extension) = &value.extension_list {
        output.extend_from_slice(&extension.xml);
    }
    for item in &value.opaque {
        output.extend_from_slice(&item.xml);
    }
    output.extend_from_slice(b"</rd2:arrayData>");
    finish(output)
}

/// Write a FeaturePropertyBags part.
pub fn write_feature_property_bags(value: &Bags) -> Result<Vec<u8>> {
    validate_bags(value)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<fpb:FeaturePropertyBags xmlns:fpb=\"");
    escape_attr(&mut output, FEATURE_BAG);
    output.extend_from_slice(b"\" xmlns:x=\"");
    escape_attr(&mut output, SPREADSHEETML);
    output.extend_from_slice(b"\" xmlns:r=\"");
    escape_attr(&mut output, RELATIONSHIPS);
    output.push(b'"');
    if let Some(count) = value.count {
        attr(&mut output, "count", &count.to_string());
    }
    output.push(b'>');
    for item in &value.bag_extensions {
        output.extend_from_slice(&item.xml);
    }
    for item in &value.values {
        write_bag(&mut output, item)?;
    }
    if let Some(extension) = &value.extension_list {
        output.extend_from_slice(&extension.xml);
    }
    for item in &value.opaque {
        output.extend_from_slice(&item.xml);
    }
    output.extend_from_slice(b"</fpb:FeaturePropertyBags>");
    finish(output)
}

/// Write a richValueRels part.
pub fn write_rich_value_rels(value: &RichValueRels) -> Result<Vec<u8>> {
    validate_rich_value_rels(value)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<rvr:richValueRels xmlns:rvr=\"");
    escape_attr(&mut output, RICH_VALUE_REL);
    output.extend_from_slice(b"\" xmlns:r=\"");
    escape_attr(&mut output, RELATIONSHIPS);
    output.extend_from_slice(b"\" xmlns:x=\"");
    escape_attr(&mut output, SPREADSHEETML);
    output.push(b'>');
    for id in &value.ids {
        output.extend_from_slice(b"<rvr:rel");
        output.extend_from_slice(b" r:id=\"");
        escape_attr(&mut output, id);
        output.extend_from_slice(b"\"/>");
    }
    if let Some(extension) = &value.extension_list {
        output.extend_from_slice(&extension.xml);
    }
    for item in &value.opaque {
        output.extend_from_slice(&item.xml);
    }
    output.extend_from_slice(b"</rvr:richValueRels>");
    finish(output)
}

/// Write one xfComplement extension element.
pub fn write_xf_complement(value: &XfComplement) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output.extend_from_slice(b"<fpb:xfComplement xmlns:fpb=\"");
    escape_attr(&mut output, FEATURE_BAG);
    attr(&mut output, "i", &value.index.to_string());
    write_opaque_children(&mut output, &value.opaque);
    output.extend_from_slice(b"</fpb:xfComplement>");
    finish(output)
}

/// Write one DXFComplement extension element.
pub fn write_dxf_complement(value: &DxfComplement) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output.extend_from_slice(b"<fpb:DXFComplement xmlns:fpb=\"");
    escape_attr(&mut output, FEATURE_BAG);
    attr(&mut output, "i", &value.index.to_string());
    write_opaque_children(&mut output, &value.opaque);
    output.extend_from_slice(b"</fpb:DXFComplement>");
    finish(output)
}

pub(crate) fn write_part(kind: Kind, document: &Document) -> Result<Vec<u8>> {
    match (kind, document) {
        (Kind::Data, Document::Data(value)) => write_data(value),
        (Kind::Structures, Document::Structures(value)) => write_structures(value),
        (Kind::Arrays, Document::Arrays(value)) => write_arrays(value),
        (Kind::Relationships, Document::Relationships(value)) => write_rich_value_rels(value),
        (Kind::FeatureBags, Document::FeatureBags(value)) => write_feature_property_bags(value),
        (
            Kind::Styles
            | Kind::SupportingData
            | Kind::SupportingStructures
            | Kind::Types
            | Kind::WebImages,
            Document::Opaque(value),
        ) => Ok(value.xml.clone()),
        _ => Err(invalid(
            "rich-values document kind does not match its part kind",
        )),
    }
}

fn write_value(output: &mut Vec<u8>, value: &RichValue) -> Result<()> {
    output.extend_from_slice(b"<rd:rv");
    attr(output, "s", &value.structure.to_string());
    if value.fallback.is_none() && value.values.is_empty() {
        return Err(invalid("rich-value requires values"));
    }
    output.push(b'>');
    if let Some(fallback) = &value.fallback {
        write_fallback(output, fallback);
    }
    for item in &value.values {
        output.extend_from_slice(b"<rd:v>");
        escape_text(output, item);
        output.extend_from_slice(b"</rd:v>");
    }
    write_opaque_children(output, &value.opaque);
    output.extend_from_slice(b"</rd:rv>");
    Ok(())
}

fn write_fallback(output: &mut Vec<u8>, value: &Fallback) {
    output.extend_from_slice(b"<rd:fb");
    attr(output, "t", value.value_type.token());
    output.push(b'>');
    escape_text(output, &value.value);
    output.extend_from_slice(b"</rd:fb>");
}

fn write_structure(output: &mut Vec<u8>, value: &Structure) -> Result<()> {
    output.extend_from_slice(b"<rd:s");
    attr(output, "t", &value.type_name);
    output.push(b'>');
    for key in &value.keys {
        output.extend_from_slice(b"<rd:k");
        attr(output, "n", &key.name);
        attr(output, "t", key.value_type.token());
        output.extend_from_slice(b"/>");
    }
    write_opaque_children(output, &value.opaque);
    output.extend_from_slice(b"</rd:s>");
    Ok(())
}

fn write_array(output: &mut Vec<u8>, value: &Array) -> Result<()> {
    output.extend_from_slice(b"<rd2:a");
    attr(output, "r", &value.rows.to_string());
    if value.columns != 1 {
        attr(output, "c", &value.columns.to_string());
    }
    output.push(b'>');
    for item in &value.values {
        write_array_value(output, item);
    }
    write_opaque_children(output, &value.opaque);
    output.extend_from_slice(b"</rd2:a>");
    Ok(())
}

fn write_array_value(output: &mut Vec<u8>, value: &ArrayValue) {
    output.extend_from_slice(b"<rd2:v");
    attr(output, "t", value.value_type.token());
    output.push(b'>');
    escape_text(output, &value.value);
    output.extend_from_slice(b"</rd2:v>");
}

fn write_bag(output: &mut Vec<u8>, value: &Bag) -> Result<()> {
    output.extend_from_slice(b"<fpb:bag");
    attr(output, "type", value.bag_type.token());
    if let Some(ext_ref) = &value.ext_ref {
        attr(output, "extRef", ext_ref);
    }
    if let Some(index) = value.bag_extension {
        attr(output, "bagExtId", &index.to_string());
    }
    if let Some(attribute) = &value.attribute {
        attr(output, "att", attribute);
    }
    if value.properties.is_empty() && value.opaque.is_empty() {
        output.extend_from_slice(b"/>");
        return Ok(());
    }
    output.push(b'>');
    for property in &value.properties {
        write_property(output, property)?;
    }
    write_opaque_children(output, &value.opaque);
    output.extend_from_slice(b"</fpb:bag>");
    Ok(())
}

fn write_property(output: &mut Vec<u8>, value: &Property) -> Result<()> {
    match value {
        Property::Array { key, values } => {
            output.extend_from_slice(b"<fpb:a");
            attr(output, "k", key);
            if values.is_empty() {
                output.extend_from_slice(b"/>");
            } else {
                output.push(b'>');
                for item in values {
                    write_property_value(output, item)?;
                }
                output.extend_from_slice(b"</fpb:a>");
            }
        },
        Property::Bag { key, index } => write_scalar(output, "bagId", key, &index.to_string()),
        Property::Integer { key, value } => write_scalar(output, "i", key, value),
        Property::Text { key, value } => write_scalar(output, "s", key, value),
        Property::Boolean { key, value } => {
            write_scalar(output, "b", key, if *value { "1" } else { "0" })
        },
        Property::Decimal { key, value } => write_scalar(output, "d", key, value),
        Property::Relationship { key, id } => write_scalar(output, "rel", key, id),
        Property::Unknown(value) => output.extend_from_slice(&value.xml),
    }
    Ok(())
}

fn write_scalar(output: &mut Vec<u8>, name: &str, key: &str, value: &str) {
    output.extend_from_slice(b"<fpb:");
    output.extend_from_slice(name.as_bytes());
    attr(output, "k", key);
    output.push(b'>');
    escape_text(output, value);
    output.extend_from_slice(b"</fpb:");
    output.extend_from_slice(name.as_bytes());
    output.push(b'>');
}

fn write_property_value(output: &mut Vec<u8>, value: &PropertyValue) -> Result<()> {
    match value {
        PropertyValue::Bag(value) => write_leaf(output, "bagId", &value.to_string()),
        PropertyValue::Integer(value) => write_leaf(output, "i", value),
        PropertyValue::Text(value) => write_leaf(output, "s", value),
        PropertyValue::Boolean(value) => write_leaf(output, "b", if *value { "1" } else { "0" }),
        PropertyValue::Decimal(value) => write_leaf(output, "d", value),
        PropertyValue::Relationship(value) => write_leaf(output, "rel", value),
        PropertyValue::Unknown(value) => output.extend_from_slice(&value.xml),
    }
    Ok(())
}

fn write_leaf(output: &mut Vec<u8>, name: &str, value: &str) {
    output.extend_from_slice(b"<fpb:");
    output.extend_from_slice(name.as_bytes());
    output.push(b'>');
    escape_text(output, value);
    output.extend_from_slice(b"</fpb:");
    output.extend_from_slice(name.as_bytes());
    output.push(b'>');
}

fn write_opaque_children(output: &mut Vec<u8>, values: &[super::super::model::Opaque]) {
    for value in values {
        output.extend_from_slice(&value.xml);
    }
}

fn finish(output: Vec<u8>) -> Result<Vec<u8>> {
    if output.len() > MAX_OUTPUT_BYTES {
        return Err(limit("serialized XML"));
    }
    Ok(output)
}
