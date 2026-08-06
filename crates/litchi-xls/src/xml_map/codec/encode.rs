//! Deterministic XML-map serializer.

use super::super::model::{DataBinding, Map, MapInfo, OpaqueXml, Schema};
use super::super::validation;
use super::parse::MAX_STREAM_BYTES;
use crate::{Error, Result};

/// Serialize a complete XML-map stream with canonical known markup.
pub fn write(value: &MapInfo) -> Result<Vec<u8>> {
    validation::validate(value)?;
    let mut output = Vec::with_capacity(512);
    push(&mut output, b"<?xml version=\"1.0\" encoding=\"utf-8\"?>")?;
    push(&mut output, b"<MapInfo")?;
    write_namespaces(&mut output, value.namespaces())?;
    attribute(
        &mut output,
        b"SelectionNamespaces",
        value.selection_namespaces(),
    )?;
    push(&mut output, b">")?;
    for schema in value.schemas() {
        write_schema(&mut output, schema)?;
    }
    for map in value.maps() {
        write_map(&mut output, map)?;
    }
    push(&mut output, b"</MapInfo>")?;
    Ok(output)
}

fn write_schema(output: &mut Vec<u8>, schema: &Schema) -> Result<()> {
    push(output, b"<Schema")?;
    attribute(output, b"ID", schema.id().as_str())?;
    if let Some(references) = schema.schema_references() {
        let value = references
            .iter()
            .map(|reference| reference.as_str())
            .collect::<Vec<_>>()
            .join(" ");
        attribute(output, b"SchemaRef", &value)?;
    }
    if let Some(namespace) = schema.namespace() {
        attribute(output, b"Namespace", namespace)?;
    }
    write_namespaces(output, schema.namespaces())?;
    push(output, b">")?;
    opaque(output, schema.payload())?;
    push(output, b"</Schema>")
}

fn write_map(output: &mut Vec<u8>, map: &Map) -> Result<()> {
    push(output, b"<Map")?;
    attribute(output, b"ID", &map.id().get().to_string())?;
    attribute(output, b"Name", map.name())?;
    attribute(output, b"RootElement", map.root_element())?;
    attribute(output, b"SchemaID", map.schema_id().as_str())?;
    boolean(
        output,
        b"ShowImportExportValidationErrors",
        map.show_import_export_validation_errors(),
    )?;
    boolean(output, b"AutoFit", map.auto_fit())?;
    boolean(output, b"Append", map.append())?;
    boolean(
        output,
        b"PreserveSortAFLayout",
        map.preserve_sort_auto_filter_layout(),
    )?;
    boolean(output, b"PreserveFormat", map.preserve_format())?;
    write_namespaces(output, map.namespaces())?;
    if let Some(binding) = map.data_binding() {
        push(output, b">")?;
        write_binding(output, binding)?;
        push(output, b"</Map>")
    } else {
        push(output, b"/>")
    }
}

fn write_binding(output: &mut Vec<u8>, binding: &DataBinding) -> Result<()> {
    push(output, b"<DataBinding")?;
    if let Some(name) = binding.data_binding_name() {
        attribute(output, b"DataBindingName", name)?;
    }
    attribute(output, b"FileBinding", binding.file_binding())?;
    if let Some(name) = binding.file_binding_name() {
        attribute(output, b"FileBindingName", name)?;
    }
    attribute(
        output,
        b"DataBindingLoadMode",
        &binding.load_mode().code().to_string(),
    )?;
    write_namespaces(output, binding.namespaces())?;
    if let Some(payload) = binding.payload() {
        push(output, b">")?;
        opaque(output, payload)?;
        push(output, b"</DataBinding>")
    } else {
        push(output, b"/>")
    }
}

fn write_namespaces(
    output: &mut Vec<u8>,
    namespaces: &[super::super::model::NamespaceDeclaration],
) -> Result<()> {
    for namespace in namespaces {
        if namespace.prefix().is_empty() {
            attribute(output, b"xmlns", namespace.uri())?;
        } else {
            let mut name = b"xmlns:".to_vec();
            name.extend_from_slice(namespace.prefix().as_bytes());
            attribute(output, &name, namespace.uri())?;
        }
    }
    Ok(())
}

fn opaque(output: &mut Vec<u8>, value: &OpaqueXml) -> Result<()> {
    push(output, value.as_bytes())
}

fn boolean(output: &mut Vec<u8>, name: &[u8], value: bool) -> Result<()> {
    attribute(output, name, if value { "true" } else { "false" })
}

fn attribute(output: &mut Vec<u8>, name: &[u8], value: &str) -> Result<()> {
    push(output, b" ")?;
    push(output, name)?;
    push(output, b"=\"")?;
    let escaped = litchi_core::xml::escape_xml(value);
    push(output, escaped.as_bytes())?;
    push(output, b"\"")
}

fn push(output: &mut Vec<u8>, bytes: &[u8]) -> Result<()> {
    let new_len = output
        .len()
        .checked_add(bytes.len())
        .ok_or_else(|| limit("serialized XML map length overflows"))?;
    if new_len > MAX_STREAM_BYTES {
        return Err(limit("serialized XML map exceeds the 16 MiB limit"));
    }
    output.extend_from_slice(bytes);
    Ok(())
}

fn limit(message: impl Into<String>) -> Error {
    Error::InvalidData(format!("XML map resource limit: {}", message.into()))
}
