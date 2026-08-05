//! Semantic web-extension and task-pane values at the XML boundary.

use super::super::model::{
    AddIn, Binding, BindingKind, Compression, Conformance, Dock, Effect, ExtKind, ExtList, Limits,
    OperationBudget, Pane, Panes, Property, Reference, Snapshot, SnapshotTarget, Store,
};
use super::super::package::fold_part_name;
use super::super::{
    ADD_IN_RELATIONSHIP, IMAGE_RELATIONSHIP_TYPE, STRICT_RELATIONSHIPS_NAMESPACE,
    TASK_PANES_NAMESPACE, TRANSITIONAL_RELATIONSHIPS_NAMESPACE, WEB_EXTENSION_NAMESPACE,
};
use super::relationship::relationship_attr;
use super::xml::{
    Node, XmlDocument, attr, element_children, ensure_consumed, is_drawingml_namespace, is_next,
    next_required, optional_bool_attr, parse_mce_xml, reject_unknown_attributes, require_name,
    required_attr,
};
use crate::{Error, Result};
use std::collections::{HashMap, HashSet};

/// Parse one MS-OWEXML web extension part after bounded MCE preprocessing.
#[cfg(test)]
pub(in crate::web) fn parse_add_in(xml: &[u8]) -> Result<AddIn> {
    parse_add_in_with(xml, &Limits::standard())
}

pub(in crate::web) fn parse_add_in_with(xml: &[u8], limits: &Limits) -> Result<AddIn> {
    let mut budget = OperationBudget::default();
    parse_add_in_with_budget(xml, limits, &mut budget)
}

pub(in crate::web) fn parse_add_in_with_budget(
    xml: &[u8],
    limits: &Limits,
    budget: &mut OperationBudget,
) -> Result<AddIn> {
    budget.charge_xml(xml.len(), limits)?;
    let document = parse_mce_xml(xml, &[WEB_EXTENSION_NAMESPACE], limits)?;
    budget.charge_strings(
        document
            .xml
            .len()
            .checked_add(document.string_bytes)
            .ok_or(Error::Limit {
                resource: "retained web extension string bytes",
                max: limits.total_string_bytes,
                actual: usize::MAX,
            })?,
        limits,
    )?;
    let root = document.root()?;
    require_name(root, WEB_EXTENSION_NAMESPACE, "webextension")?;
    reject_unknown_attributes(root, &[("", "id"), ("", "frozen")])?;

    let id = required_attr(root, "", "id")?.to_owned();
    let frozen = optional_bool_attr(root, "", "frozen")?.unwrap_or(false);
    let children = element_children(root);
    let mut position = 0;

    let reference_node = next_required(
        &children,
        &mut position,
        WEB_EXTENSION_NAMESPACE,
        "reference",
    )?;
    let reference = parse_store_reference(reference_node, &document)?;

    let alternate_references = if is_next(
        &children,
        position,
        WEB_EXTENSION_NAMESPACE,
        "alternateReferences",
    ) {
        let node = children[position];
        position += 1;
        reject_unknown_attributes(node, &[])?;
        let refs = element_children(node);
        enforce_count_with("alternate reference", refs.len(), limits)?;
        refs.into_iter()
            .map(|child| {
                require_name(child, WEB_EXTENSION_NAMESPACE, "reference")?;
                parse_store_reference(child, &document)
            })
            .collect::<Result<Vec<_>>>()?
    } else {
        Vec::new()
    };

    let properties_node = next_required(
        &children,
        &mut position,
        WEB_EXTENSION_NAMESPACE,
        "properties",
    )?;
    reject_unknown_attributes(properties_node, &[])?;
    let property_nodes = element_children(properties_node);
    enforce_count_with("property", property_nodes.len(), limits)?;
    let properties = property_nodes
        .into_iter()
        .map(parse_property)
        .collect::<Result<Vec<_>>>()?;

    let bindings_node = next_required(
        &children,
        &mut position,
        WEB_EXTENSION_NAMESPACE,
        "bindings",
    )?;
    reject_unknown_attributes(bindings_node, &[])?;
    let binding_nodes = element_children(bindings_node);
    enforce_count_with("binding", binding_nodes.len(), limits)?;
    let bindings = binding_nodes
        .into_iter()
        .map(|node| parse_binding(node, &document))
        .collect::<Result<Vec<_>>>()?;

    let snapshot = if is_next(&children, position, WEB_EXTENSION_NAMESPACE, "snapshot") {
        let node = children[position];
        position += 1;
        reject_unknown_attributes(
            node,
            &[
                ("", "cstate"),
                (TRANSITIONAL_RELATIONSHIPS_NAMESPACE, "embed"),
                (TRANSITIONAL_RELATIONSHIPS_NAMESPACE, "link"),
                (STRICT_RELATIONSHIPS_NAMESPACE, "embed"),
                (STRICT_RELATIONSHIPS_NAMESPACE, "link"),
            ],
        )?;
        let embedded_relationship_id = relationship_attr(node, "embed")?.map(str::to_owned);
        let linked_relationship_id = relationship_attr(node, "link")?.map(str::to_owned);
        let compression_state = attr(node, "", "cstate")
            .map(Compression::parse)
            .transpose()?;
        let snapshot_children = element_children(node);
        enforce_count_with("snapshot effect", snapshot_children.len(), limits)?;
        let mut effects = Vec::with_capacity(snapshot_children.len());
        let mut extension_list = None;
        for (index, child) in snapshot_children.iter().enumerate() {
            if is_drawingml_namespace(&child.namespace) && child.local_name == "extLst" {
                if index + 1 != snapshot_children.len() {
                    return invalid("snapshot extLst must be the final child".into());
                }
                extension_list = Some(ExtList::from_node(child, &document)?);
                continue;
            }
            effects.push(Effect::from_node(child)?);
        }
        Some(Snapshot {
            embedded_relationship_id,
            linked_relationship_id,
            compression_state,
            effects,
            extension_list,
        })
    } else {
        None
    };

    let extension_list = if is_next(&children, position, WEB_EXTENSION_NAMESPACE, "extLst") {
        let value = ExtList::from_node(children[position], &document)?;
        position += 1;
        Some(value)
    } else {
        None
    };
    ensure_consumed(&children, position, "webextension")?;

    Ok(AddIn {
        id,
        frozen,
        reference,
        alternate_references,
        properties,
        bindings,
        snapshot,
        extension_list,
    })
}

/// Parse task-pane metadata without resolving its web-extension relationships.
#[cfg(test)]
pub(in crate::web) fn parse_panes(xml: &[u8]) -> Result<Vec<ParsedPane>> {
    parse_panes_with(xml, &Limits::standard())
}

pub(in crate::web) fn parse_panes_with(xml: &[u8], limits: &Limits) -> Result<Vec<ParsedPane>> {
    let mut budget = OperationBudget::default();
    parse_panes_with_budget(xml, limits, &mut budget)
}

pub(in crate::web) fn parse_panes_with_budget(
    xml: &[u8],
    limits: &Limits,
    budget: &mut OperationBudget,
) -> Result<Vec<ParsedPane>> {
    budget.charge_xml(xml.len(), limits)?;
    let document = parse_mce_xml(
        xml,
        &[TASK_PANES_NAMESPACE, WEB_EXTENSION_NAMESPACE],
        limits,
    )?;
    budget.charge_strings(
        document
            .xml
            .len()
            .checked_add(document.string_bytes)
            .ok_or(Error::Limit {
                resource: "retained web extension string bytes",
                max: limits.total_string_bytes,
                actual: usize::MAX,
            })?,
        limits,
    )?;
    let root = document.root()?;
    require_name(root, TASK_PANES_NAMESPACE, "taskpanes")?;
    reject_unknown_attributes(root, &[])?;
    let children = element_children(root);
    enforce_count_with("task pane", children.len(), limits)?;
    children
        .into_iter()
        .map(|node| parse_task_pane(node, &document))
        .collect()
}

#[derive(Debug, Clone, PartialEq)]
pub(in crate::web) struct ParsedPane {
    pub(in crate::web) dock_state: Dock,
    pub(in crate::web) visible: bool,
    pub(in crate::web) width: f64,
    pub(in crate::web) row: u32,
    pub(in crate::web) locked: bool,
    pub(in crate::web) relationship_id: String,
    pub(in crate::web) extension_list: Option<ExtList>,
}

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

pub(in crate::web) fn parse_task_pane(node: &Node, document: &XmlDocument) -> Result<ParsedPane> {
    require_name(node, TASK_PANES_NAMESPACE, "taskpane")?;
    reject_unknown_attributes(
        node,
        &[
            ("", "dockstate"),
            ("", "visibility"),
            ("", "width"),
            ("", "row"),
            ("", "locked"),
        ],
    )?;
    let dock_state = Dock::parse(required_attr(node, "", "dockstate")?)?;
    let visible = parse_bool(required_attr(node, "", "visibility")?)?;
    let width = required_attr(node, "", "width")?
        .parse::<f64>()
        .map_err(|_| Error::Invalid("invalid task-pane width".into()))?;
    if !width.is_finite() {
        return invalid("task-pane width must be finite".into());
    }
    let row = required_attr(node, "", "row")?
        .parse::<u32>()
        .map_err(|_| Error::Invalid("invalid task-pane row".into()))?;
    let locked = optional_bool_attr(node, "", "locked")?.unwrap_or(false);
    let children = element_children(node);
    if children.is_empty() {
        return invalid("taskpane requires webextensionref".into());
    }
    let reference = children[0];
    require_name(reference, TASK_PANES_NAMESPACE, "webextensionref")?;
    reject_unknown_attributes(
        reference,
        &[
            (TRANSITIONAL_RELATIONSHIPS_NAMESPACE, "id"),
            (STRICT_RELATIONSHIPS_NAMESPACE, "id"),
        ],
    )?;
    let relationship_id = relationship_attr(reference, "id")?
        .ok_or_else(|| Error::Invalid("webextensionref requires r:id".into()))?
        .to_owned();
    if children.len() > 2
        || (children.len() == 2
            && (children[1].namespace != TASK_PANES_NAMESPACE
                || children[1].local_name != "extLst"))
    {
        return invalid("unexpected taskpane child or child order".into());
    }
    let extension_list = children
        .get(1)
        .map(|node| ExtList::from_node(node, document))
        .transpose()?;
    Ok(ParsedPane {
        dock_state,
        visible,
        width,
        row,
        locked,
        relationship_id,
        extension_list,
    })
}

pub(in crate::web) fn parse_store_reference(
    node: &Node,
    document: &XmlDocument,
) -> Result<Reference> {
    require_name(node, WEB_EXTENSION_NAMESPACE, "reference")?;
    reject_unknown_attributes(
        node,
        &[
            ("", "id"),
            ("", "version"),
            ("", "store"),
            ("", "storeType"),
        ],
    )?;
    let children = element_children(node);
    if children.len() > 1
        || children.first().is_some_and(|child| {
            child.namespace != WEB_EXTENSION_NAMESPACE || child.local_name != "extLst"
        })
    {
        return invalid("reference permits only one trailing extLst".into());
    }
    let reference = Reference {
        id: required_attr(node, "", "id")?.to_owned(),
        version: required_attr(node, "", "version")?.to_owned(),
        location: attr(node, "", "store").map(str::to_owned),
        store: attr(node, "", "storeType")
            .map(Store::parse)
            .transpose()?
            .unwrap_or_default(),
        extension_list: children
            .first()
            .map(|node| ExtList::from_node(node, document))
            .transpose()?,
    };
    validate_store_reference(&reference)?;
    Ok(reference)
}

pub(in crate::web) fn parse_property(node: &Node) -> Result<Property> {
    require_name(node, WEB_EXTENSION_NAMESPACE, "property")?;
    reject_unknown_attributes(node, &[("", "name"), ("", "value")])?;
    if !element_children(node).is_empty() {
        return invalid("web extension property must be empty".into());
    }
    Ok(Property {
        name: required_attr(node, "", "name")?.to_owned(),
        value: required_attr(node, "", "value")?.to_owned(),
    })
}

pub(in crate::web) fn parse_binding(node: &Node, document: &XmlDocument) -> Result<Binding> {
    require_name(node, WEB_EXTENSION_NAMESPACE, "binding")?;
    reject_unknown_attributes(node, &[("", "id"), ("", "type"), ("", "appref")])?;
    let children = element_children(node);
    if children.len() > 1
        || children.first().is_some_and(|child| {
            child.namespace != WEB_EXTENSION_NAMESPACE || child.local_name != "extLst"
        })
    {
        return invalid("binding permits only one trailing extLst".into());
    }
    Ok(Binding {
        id: required_attr(node, "", "id")?.to_owned(),
        kind: BindingKind::parse(required_attr(node, "", "type")?)?,
        app_ref: required_attr(node, "", "appref")?.to_owned(),
        extension_list: children
            .first()
            .map(|node| ExtList::from_node(node, document))
            .transpose()?,
    })
}

pub(in crate::web) fn validate_model(extension: &AddIn) -> Result<()> {
    validate_model_with(extension, &Limits::standard())
}

pub(in crate::web) fn validate_model_with(extension: &AddIn, limits: &Limits) -> Result<()> {
    require_nonempty("web extension id", &extension.id)?;
    validate_store_reference(&extension.reference)?;
    enforce_count_with(
        "alternate reference",
        extension.alternate_references.len(),
        limits,
    )?;
    enforce_count_with("property", extension.properties.len(), limits)?;
    enforce_count_with("binding", extension.bindings.len(), limits)?;
    let mut reference_ids = HashSet::new();
    reference_ids.insert(extension.reference.id.as_str());
    for reference in &extension.alternate_references {
        validate_store_reference(reference)?;
        if !reference_ids.insert(reference.id.as_str()) {
            return invalid(format!("duplicate reference id '{}'", reference.id));
        }
    }
    let mut property_names = HashSet::new();
    for property in &extension.properties {
        require_nonempty("property name", &property.name)?;
        if !property_names.insert(property.name.as_str()) {
            return invalid(format!("duplicate property name '{}'", property.name));
        }
    }
    let mut binding_ids = HashSet::new();
    let mut binding_app_refs = HashSet::new();
    for binding in &extension.bindings {
        validate_binding(binding)?;
        if !binding_ids.insert(binding.id.as_str()) {
            return invalid(format!("duplicate binding id '{}'", binding.id));
        }
        if !binding_app_refs.insert(binding.app_ref.as_str()) {
            return invalid(format!("duplicate binding appRef '{}'", binding.app_ref));
        }
    }
    if let Some(snapshot) = &extension.snapshot {
        enforce_count_with("snapshot effect", snapshot.effects.len(), limits)?;
        for effect in &snapshot.effects {
            let reparsed = Effect::from_xml(effect.xml.as_bytes())?;
            if reparsed.kind != effect.kind {
                return invalid("snapshot effect kind does not match its XML root".into());
            }
        }
        validate_extension_list(
            snapshot.extension_list.as_ref(),
            &[ExtKind::DrawingMl, ExtKind::StrictDrawingMl],
        )?;
    }
    validate_extension_list(extension.extension_list.as_ref(), &[ExtKind::AddIn])?;
    Ok(())
}

pub(in crate::web) fn validate_binding(binding: &Binding) -> Result<()> {
    require_nonempty("binding id", &binding.id)?;
    require_nonempty("binding type", binding.kind.as_str())?;
    require_nonempty("binding appref", &binding.app_ref)?;
    validate_extension_list(binding.extension_list.as_ref(), &[ExtKind::AddIn])
}

pub(in crate::web) fn validate_store_reference(reference: &Reference) -> Result<()> {
    require_nonempty("reference id", &reference.id)?;
    require_nonempty("reference version", &reference.version)?;
    if let Some(location) = &reference.location {
        require_nonempty("reference location", location)?;
    } else if reference.store == Store::FileSystem {
        return invalid("FileSystem reference requires a non-empty location".into());
    }
    validate_extension_list(reference.extension_list.as_ref(), &[ExtKind::AddIn])
}

pub(in crate::web) fn validate_task_pane(pane: &Pane) -> Result<()> {
    validate_task_pane_with(pane, &Limits::standard())
}

pub(in crate::web) fn validate_task_pane_with(pane: &Pane, limits: &Limits) -> Result<()> {
    require_nonempty("dock state", pane.dock_state.as_str())?;
    require_nonempty("task-pane relationship id", &pane.relationship_id)?;
    if !pane.width.is_finite() || pane.width <= 0.0 {
        return invalid("task-pane width must be finite and positive".into());
    }
    validate_extension_list(pane.extension_list.as_ref(), &[ExtKind::TaskPane])?;
    validate_model_with(&pane.add_in, limits)?;
    validate_snapshot_resources_with(pane, limits)
}

pub(in crate::web) fn validate_panes(task_panes: &Panes, limits: &Limits) -> Result<()> {
    enforce_count_with("task pane", task_panes.panes.len(), limits)?;
    let mut relationship_ids = HashSet::new();
    let mut extension_ids = HashSet::new();
    let mut total_snapshot_bytes = 0usize;
    let mut snapshot_names = HashSet::new();
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
        for resource in &pane.snapshot_resources {
            if let SnapshotTarget::Internal {
                part_name, data, ..
            } = &resource.target
            {
                if data.len() > limits.image_bytes {
                    return limit(
                        "web extension snapshot bytes",
                        limits.image_bytes,
                        data.len(),
                    );
                }
                if snapshot_names.insert(fold_part_name(part_name)) {
                    total_snapshot_bytes =
                        total_snapshot_bytes
                            .checked_add(data.len())
                            .ok_or(Error::Limit {
                                resource: "aggregate web extension snapshot bytes",
                                max: limits.total_image_bytes,
                                actual: usize::MAX,
                            })?;
                    if total_snapshot_bytes > limits.total_image_bytes {
                        return limit(
                            "aggregate web extension snapshot bytes",
                            limits.total_image_bytes,
                            total_snapshot_bytes,
                        );
                    }
                }
            }
        }
    }
    Ok(())
}

pub(in crate::web) fn charge_authored_metadata(
    task_panes: &Panes,
    budget: &mut OperationBudget,
    limits: &Limits,
) -> Result<()> {
    let generated_names = task_panes
        .panes
        .len()
        .checked_add(1)
        .and_then(|count| count.checked_mul(128))
        .ok_or(Error::Limit {
            resource: "authored web extension package metadata bytes",
            max: limits.total_string_bytes,
            actual: usize::MAX,
        })?;
    budget.charge_metadata(generated_names, 2, limits)?;
    for pane in &task_panes.panes {
        let pane_bytes = pane
            .relationship_id
            .len()
            .checked_add(ADD_IN_RELATIONSHIP.len())
            .ok_or(Error::Limit {
                resource: "authored web extension package metadata bytes",
                max: limits.total_string_bytes,
                actual: usize::MAX,
            })?;
        budget.charge_metadata(pane_bytes, 2, limits)?;
        for resource in &pane.snapshot_resources {
            let target_bytes = match &resource.target {
                SnapshotTarget::Internal {
                    part_name,
                    content_type,
                    ..
                } => part_name.as_str().len().checked_add(content_type.len()),
                SnapshotTarget::External { target } => Some(target.len()),
            }
            .and_then(|bytes| bytes.checked_add(resource.relationship_id.len()))
            .and_then(|bytes| bytes.checked_add(IMAGE_RELATIONSHIP_TYPE.len()))
            .ok_or(Error::Limit {
                resource: "authored web extension package metadata bytes",
                max: limits.total_string_bytes,
                actual: usize::MAX,
            })?;
            budget.charge_metadata(target_bytes, 3, limits)?;
        }
    }
    Ok(())
}

pub(in crate::web) fn add_xml_budget(
    total: &mut usize,
    bytes: usize,
    limits: &Limits,
) -> Result<()> {
    *total = total.checked_add(bytes).ok_or(Error::Limit {
        resource: "authored web extension XML bytes",
        max: limits.xml_bytes,
        actual: usize::MAX,
    })?;
    if *total > limits.xml_bytes {
        return limit("authored web extension XML bytes", limits.xml_bytes, *total);
    }
    Ok(())
}

pub(in crate::web) fn add_escaped_xml_budget(
    total: &mut usize,
    value: &str,
    limits: &Limits,
) -> Result<()> {
    let bytes = value.len().checked_mul(6).ok_or(Error::Limit {
        resource: "authored web extension XML bytes",
        max: limits.xml_bytes,
        actual: usize::MAX,
    })?;
    add_xml_budget(total, bytes, limits)
}

pub(in crate::web) fn add_reference_budget(
    total: &mut usize,
    reference: &Reference,
    limits: &Limits,
) -> Result<()> {
    add_xml_budget(total, 128, limits)?;
    add_escaped_xml_budget(total, &reference.id, limits)?;
    add_escaped_xml_budget(total, &reference.version, limits)?;
    if let Some(location) = &reference.location {
        add_escaped_xml_budget(total, location, limits)?;
    }
    if let Some(extension_list) = &reference.extension_list {
        add_xml_budget(total, extension_list.xml.len(), limits)?;
    }
    Ok(())
}

pub(in crate::web) fn validate_add_in_budget(extension: &AddIn, limits: &Limits) -> Result<()> {
    let mut total = 512usize;
    add_escaped_xml_budget(&mut total, &extension.id, limits)?;
    add_reference_budget(&mut total, &extension.reference, limits)?;
    for reference in &extension.alternate_references {
        add_reference_budget(&mut total, reference, limits)?;
    }
    for property in &extension.properties {
        add_xml_budget(&mut total, 64, limits)?;
        add_escaped_xml_budget(&mut total, &property.name, limits)?;
        add_escaped_xml_budget(&mut total, &property.value, limits)?;
    }
    for binding in &extension.bindings {
        add_xml_budget(&mut total, 96, limits)?;
        add_escaped_xml_budget(&mut total, &binding.id, limits)?;
        add_escaped_xml_budget(&mut total, binding.kind.as_str(), limits)?;
        add_escaped_xml_budget(&mut total, &binding.app_ref, limits)?;
        if let Some(extension_list) = &binding.extension_list {
            add_xml_budget(&mut total, extension_list.xml.len(), limits)?;
        }
    }
    if let Some(snapshot) = &extension.snapshot {
        add_xml_budget(&mut total, 160, limits)?;
        for effect in &snapshot.effects {
            add_xml_budget(&mut total, effect.xml.len(), limits)?;
        }
        if let Some(extension_list) = &snapshot.extension_list {
            add_xml_budget(&mut total, extension_list.xml.len(), limits)?;
        }
    }
    if let Some(extension_list) = &extension.extension_list {
        add_xml_budget(&mut total, extension_list.xml.len(), limits)?;
    }
    Ok(())
}

pub(in crate::web) fn validate_panes_budget(task_panes: &Panes, limits: &Limits) -> Result<()> {
    let mut total = 384usize;
    for pane in &task_panes.panes {
        add_xml_budget(&mut total, 192, limits)?;
        add_escaped_xml_budget(&mut total, pane.dock_state.as_str(), limits)?;
        add_escaped_xml_budget(&mut total, &pane.relationship_id, limits)?;
        if let Some(extension_list) = &pane.extension_list {
            add_xml_budget(&mut total, extension_list.xml.len(), limits)?;
        }
    }
    Ok(())
}

pub(in crate::web) fn validate_extension_list(
    extension_list: Option<&ExtList>,
    allowed: &[ExtKind],
) -> Result<()> {
    let Some(extension_list) = extension_list else {
        return Ok(());
    };
    if !allowed.contains(&extension_list.kind) {
        return invalid(format!(
            "extLst namespace '{}' is not valid at this location",
            extension_list.kind.namespace()
        ));
    }
    let reparsed = ExtList::from_xml(extension_list.as_xml())?;
    if reparsed != *extension_list {
        return invalid("extLst fragment is not a stable self-contained XML tree".into());
    }
    Ok(())
}

pub(in crate::web) fn validate_snapshot_resources_with(pane: &Pane, limits: &Limits) -> Result<()> {
    let mut expected = HashMap::new();
    if let Some(snapshot) = &pane.add_in.snapshot {
        if let Some(id) = snapshot.embedded_relationship_id.as_deref() {
            require_nonempty("embedded snapshot relationship ID", id)?;
            expected.insert(id, false);
        }
        if let Some(id) = snapshot.linked_relationship_id.as_deref() {
            require_nonempty("linked snapshot relationship ID", id)?;
            if expected.insert(id, true).is_some() {
                return invalid("snapshot embed and link IDs must differ".into());
            }
        }
    }
    if expected.len() != pane.snapshot_resources.len() {
        return invalid("snapshot relationship and resource counts differ".into());
    }
    let mut resource_ids = HashSet::new();
    for resource in &pane.snapshot_resources {
        require_nonempty(
            "snapshot resource relationship ID",
            &resource.relationship_id,
        )?;
        if !resource_ids.insert(resource.relationship_id.as_str()) {
            return invalid(format!(
                "duplicate snapshot resource relationship ID '{}'",
                resource.relationship_id
            ));
        }
        let Some(linked) = expected.get(resource.relationship_id.as_str()) else {
            return invalid(format!(
                "snapshot resource '{}' is not referenced by the web extension",
                resource.relationship_id
            ));
        };
        match &resource.target {
            SnapshotTarget::Internal {
                part_name,
                content_type,
                data,
            } => {
                if part_name.as_str() == "/" {
                    return invalid("snapshot image cannot target the package root".into());
                }
                validate_image_content_type(content_type)?;
                if data.len() > limits.image_bytes {
                    return limit(
                        "web extension snapshot bytes",
                        limits.image_bytes,
                        data.len(),
                    );
                }
            },
            SnapshotTarget::External { target } => {
                if !*linked {
                    return invalid(format!(
                        "embedded snapshot resource '{}' cannot be external",
                        resource.relationship_id
                    ));
                }
                validate_external_uri_reference(target)?;
            },
        }
    }
    Ok(())
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

pub(in crate::web) fn format_f64(value: f64) -> String {
    let mut buffer = ryu::Buffer::new();
    buffer.format_finite(value).to_owned()
}

pub(in crate::web) fn escape_attr(out: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(character),
        }
    }
}

pub(in crate::web) fn require_nonempty(label: &str, value: &str) -> Result<()> {
    if value.is_empty() {
        invalid(format!("{label} must not be empty"))
    } else {
        Ok(())
    }
}

pub(in crate::web) fn validate_image_content_type(value: &str) -> Result<()> {
    if value.len() > 255 || value.bytes().any(|byte| byte.is_ascii_control()) {
        return invalid(format!("invalid snapshot image content type '{value}'"));
    }
    let Some((top_level, subtype)) = value.split_once('/') else {
        return invalid(format!("invalid snapshot image content type '{value}'"));
    };
    if !top_level.eq_ignore_ascii_case("image")
        || subtype.is_empty()
        || subtype.contains('/')
        || !top_level.bytes().all(is_mime_token_byte)
        || !subtype.bytes().all(is_mime_token_byte)
    {
        return invalid(format!("invalid snapshot image content type '{value}'"));
    }
    Ok(())
}

pub(in crate::web) fn is_mime_token_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric()
        || matches!(
            byte,
            b'!' | b'#'
                | b'$'
                | b'%'
                | b'&'
                | b'\''
                | b'*'
                | b'+'
                | b'-'
                | b'.'
                | b'^'
                | b'_'
                | b'`'
                | b'|'
                | b'~'
        )
}

pub(in crate::web) fn validate_external_uri_reference(value: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > 32 * 1024
        || value.chars().any(char::is_whitespace)
        || value.chars().any(char::is_control)
        || value.contains('\\')
    {
        return invalid("external snapshot target is not a valid URI-reference".into());
    }
    let bytes = value.as_bytes();
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] == b'%' {
            if bytes
                .get(index + 1..index + 3)
                .is_none_or(|encoded| !encoded.iter().all(u8::is_ascii_hexdigit))
            {
                return invalid("external snapshot target has invalid percent encoding".into());
            }
            index += 3;
        } else {
            index += 1;
        }
    }
    let base = url::Url::parse("https://litchi.invalid/")
        .map_err(|error| Error::Uri(error.to_string()))?;
    url::Url::options()
        .base_url(Some(&base))
        .parse(value)
        .map_err(|error| Error::Uri(format!("invalid external snapshot URI-reference: {error}")))?;
    Ok(())
}

pub(in crate::web) fn enforce_count_with(
    label: &'static str,
    count: usize,
    limits: &Limits,
) -> Result<()> {
    if count > limits.items {
        limit(label, limits.items, count)
    } else {
        Ok(())
    }
}

pub(in crate::web) fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("invalid XML boolean '{value}'")),
    }
}

pub(in crate::web) fn invalid<T>(message: String) -> Result<T> {
    Err(Error::Invalid(message))
}

pub(in crate::web) fn limit<T>(resource: &'static str, max: usize, actual: usize) -> Result<T> {
    Err(Error::Limit {
        resource,
        max,
        actual,
    })
}
