use super::super::super::model::{AddIn, Conformance, ExtKind, ExtList, Limits, Panes, Reference};
use super::super::super::validation::{
    validate_add_in_budget, validate_model_with, validate_panes, validate_panes_budget,
    validate_task_pane_with,
};
use super::super::super::{TASK_PANES_NAMESPACE, WEB_EXTENSION_NAMESPACE};
use super::super::xml::{element_children, parse_mce_xml};
use super::parser::{parse_add_in_with, parse_panes_with};
use super::support::{escape_attr, format_f64, invalid, limit};
use crate::{Error, Result};
use std::collections::HashSet;

#[cfg(test)]
pub(in crate::web) fn write_add_in(extension: &AddIn, conformance: Conformance) -> Result<Vec<u8>> {
    write_add_in_with(extension, conformance, &Limits::standard())
}

pub(in crate::web) fn write_add_in_with(
    extension: &AddIn,
    conformance: Conformance,
    limits: &Limits,
) -> Result<Vec<u8>> {
    validate_model_with(extension, limits)?;
    validate_add_in_budget(extension, limits)?;
    let mut out = String::with_capacity(1024);
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    out.push_str("<we:webextension xmlns:we=\"");
    escape_attr(&mut out, WEB_EXTENSION_NAMESPACE);
    out.push_str("\" xmlns:r=\"");
    escape_attr(&mut out, conformance.relationships_namespace());
    out.push_str("\" id=\"");
    escape_attr(&mut out, &extension.id);
    out.push_str("\" frozen=\"");
    out.push_str(if extension.frozen { "true" } else { "false" });
    out.push_str("\">");
    write_store_reference(&mut out, "reference", &extension.reference);
    if !extension.alternate_references.is_empty() {
        out.push_str("<we:alternateReferences>");
        for reference in &extension.alternate_references {
            write_store_reference(&mut out, "reference", reference);
        }
        out.push_str("</we:alternateReferences>");
    }
    out.push_str("<we:properties>");
    for property in &extension.properties {
        out.push_str("<we:property name=\"");
        escape_attr(&mut out, &property.name);
        out.push_str("\" value=\"");
        escape_attr(&mut out, &property.value);
        out.push_str("\"/>");
    }
    out.push_str("</we:properties><we:bindings>");
    for binding in &extension.bindings {
        out.push_str("<we:binding id=\"");
        escape_attr(&mut out, &binding.id);
        out.push_str("\" type=\"");
        escape_attr(&mut out, binding.kind.as_str());
        out.push_str("\" appref=\"");
        escape_attr(&mut out, &binding.app_ref);
        if let Some(extension_list) = &binding.extension_list {
            out.push_str("\">");
            out.push_str(extension_list.xml());
            out.push_str("</we:binding>");
        } else {
            out.push_str("\"/>");
        }
    }
    out.push_str("</we:bindings>");
    if let Some(snapshot) = &extension.snapshot {
        out.push_str("<we:snapshot");
        if let Some(id) = &snapshot.embedded_relationship_id {
            out.push_str(" r:embed=\"");
            escape_attr(&mut out, id);
            out.push('"');
        }
        if let Some(id) = &snapshot.linked_relationship_id {
            out.push_str(" r:link=\"");
            escape_attr(&mut out, id);
            out.push('"');
        }
        if let Some(compression_state) = snapshot.compression_state {
            out.push_str(" cstate=\"");
            out.push_str(compression_state.as_str());
            out.push('"');
        }
        if snapshot.effects.is_empty() && snapshot.extension_list.is_none() {
            out.push_str("/>");
        } else {
            out.push('>');
            for effect in &snapshot.effects {
                out.push_str(effect.xml());
            }
            if let Some(extension_list) = &snapshot.extension_list {
                out.push_str(extension_list.xml());
            }
            out.push_str("</we:snapshot>");
        }
    }
    if let Some(extension_list) = &extension.extension_list {
        out.push_str(extension_list.xml());
    }
    out.push_str("</we:webextension>");
    let output = out.into_bytes();
    if output.len() > limits.xml_bytes {
        return limit("web extension XML bytes", limits.xml_bytes, output.len());
    }
    parse_add_in_with(&output, limits)?;
    Ok(output)
}

/// Replace an existing add-in `we:extLst` using offsets from the retained
/// source XML. MCE processing is rejected when it changes bytes, because
/// effective-tree offsets cannot be published as source offsets.
pub(in crate::web) fn splice_add_in_extension_list_with(
    source: &[u8],
    extension_list: &ExtList,
    limits: &Limits,
) -> Result<Vec<u8>> {
    if extension_list.kind() != ExtKind::AddIn {
        return invalid("add-in source splice requires an add-in extLst".into());
    }
    if source.len() > limits.xml_bytes {
        return limit("web extension XML bytes", limits.xml_bytes, source.len());
    }
    let document = parse_mce_xml(source, &[WEB_EXTENSION_NAMESPACE], limits)?;
    if document.xml.as_slice() != source {
        return invalid(
            "cannot source-splice add-in XML containing markup-compatibility alternatives".into(),
        );
    }
    let root = document.root()?;
    if root.namespace != WEB_EXTENSION_NAMESPACE || root.local_name != "webextension" {
        return invalid("add-in source splice requires a webextension root".into());
    }
    let extensions = element_children(root)
        .into_iter()
        .filter(|node| node.namespace == WEB_EXTENSION_NAMESPACE && node.local_name == "extLst")
        .collect::<Vec<_>>();
    let [extension] = extensions.as_slice() else {
        return invalid(
            "cannot source-splice add-in XML without exactly one retained extLst".into(),
        );
    };
    let fragment = extension
        .raw_fragment
        .as_ref()
        .ok_or_else(|| Error::Invalid("add-in extLst has no retained source span".into()))?;
    let replacement = extension_list.xml().as_bytes();
    let output_len = source
        .len()
        .checked_sub(fragment.end - fragment.start)
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or(Error::Limit {
            resource: "web extension XML bytes",
            max: limits.xml_bytes,
            actual: usize::MAX,
        })?;
    if output_len > limits.xml_bytes {
        return limit("web extension XML bytes", limits.xml_bytes, output_len);
    }
    let mut output = Vec::new();
    output
        .try_reserve(output_len)
        .map_err(|source| Error::Allocation {
            resource: "source-spliced web extension XML",
            source,
        })?;
    output.extend_from_slice(&source[..fragment.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[fragment.end..]);
    parse_add_in_with(&output, limits)?;
    Ok(output)
}

/// Deterministically serialize task-pane metadata and relationship IDs.
#[cfg(test)]
pub(in crate::web) fn write_panes(task_panes: &Panes, conformance: Conformance) -> Result<Vec<u8>> {
    write_panes_with(task_panes, conformance, &Limits::standard())
}

pub(in crate::web) fn write_panes_with(
    task_panes: &Panes,
    conformance: Conformance,
    limits: &Limits,
) -> Result<Vec<u8>> {
    validate_panes(task_panes, limits)?;
    validate_panes_budget(task_panes, limits)?;
    let mut relationship_ids = HashSet::new();
    let mut extension_ids = HashSet::new();
    let mut out = String::with_capacity(512 + task_panes.panes.len() * 160);
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    out.push_str("<wetp:taskpanes xmlns:wetp=\"");
    escape_attr(&mut out, TASK_PANES_NAMESPACE);
    out.push_str("\" xmlns:r=\"");
    escape_attr(&mut out, conformance.relationships_namespace());
    out.push_str("\">");
    for pane in &task_panes.panes {
        validate_task_pane_with(pane, limits)?;
        if !relationship_ids.insert(pane.relationship_id.as_str()) {
            return invalid(format!(
                "duplicate task-pane relationship ID '{}'",
                pane.relationship_id
            ));
        }
        if !extension_ids.insert(pane.add_in.id.as_str()) {
            return invalid(format!(
                "duplicate web extension instance ID '{}'",
                pane.add_in.id
            ));
        }
        out.push_str("<wetp:taskpane dockstate=\"");
        escape_attr(&mut out, pane.dock_state.as_str());
        out.push_str("\" visibility=\"");
        out.push_str(if pane.visible { "true" } else { "false" });
        out.push_str("\" width=\"");
        out.push_str(&format_f64(pane.width));
        out.push_str("\" row=\"");
        out.push_str(&pane.row.to_string());
        out.push_str("\" locked=\"");
        out.push_str(if pane.locked { "true" } else { "false" });
        out.push_str("\"><wetp:webextensionref r:id=\"");
        escape_attr(&mut out, &pane.relationship_id);
        out.push_str("\"/>");
        if let Some(extension_list) = &pane.extension_list {
            out.push_str(extension_list.xml());
        }
        out.push_str("</wetp:taskpane>");
    }
    out.push_str("</wetp:taskpanes>");
    let output = out.into_bytes();
    if output.len() > limits.xml_bytes {
        return limit("task-pane XML bytes", limits.xml_bytes, output.len());
    }
    parse_panes_with(&output, limits)?;
    Ok(output)
}

pub(in crate::web) fn write_store_reference(
    out: &mut String,
    element: &str,
    reference: &Reference,
) {
    out.push_str("<we:");
    out.push_str(element);
    out.push_str(" id=\"");
    escape_attr(out, &reference.id);
    out.push_str("\" version=\"");
    escape_attr(out, &reference.version);
    if let Some(store) = &reference.location {
        out.push_str("\" store=\"");
        escape_attr(out, store);
    }
    out.push_str("\" storeType=\"");
    out.push_str(reference.store.as_str());
    if let Some(extension_list) = &reference.extension_list {
        out.push_str("\">");
        out.push_str(extension_list.xml());
        out.push_str("</we:");
        out.push_str(element);
        out.push('>');
    } else {
        out.push_str("\"/>");
    }
}
