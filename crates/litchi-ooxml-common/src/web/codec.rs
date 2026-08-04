//! Bounded MS-OWEXML and DrawingML XML parsing and deterministic writing.

use super::model::*;
use super::*;

/// Parse one MS-OWEXML web extension part after bounded MCE preprocessing.
#[cfg(test)]
pub(super) fn parse_add_in(xml: &[u8]) -> Result<AddIn> {
    parse_add_in_with(xml, &Limits::standard())
}

pub(super) fn parse_add_in_with(xml: &[u8], limits: &Limits) -> Result<AddIn> {
    let mut budget = OperationBudget::default();
    parse_add_in_with_budget(xml, limits, &mut budget)
}

pub(super) fn parse_add_in_with_budget(
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
pub(super) fn parse_panes(xml: &[u8]) -> Result<Vec<ParsedPane>> {
    parse_panes_with(xml, &Limits::standard())
}

pub(super) fn parse_panes_with(xml: &[u8], limits: &Limits) -> Result<Vec<ParsedPane>> {
    let mut budget = OperationBudget::default();
    parse_panes_with_budget(xml, limits, &mut budget)
}

pub(super) fn parse_panes_with_budget(
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
pub(super) struct ParsedPane {
    pub(super) dock_state: Dock,
    pub(super) visible: bool,
    pub(super) width: f64,
    pub(super) row: u32,
    pub(super) locked: bool,
    pub(super) relationship_id: String,
    pub(super) extension_list: Option<ExtList>,
}

#[cfg(test)]
pub(super) fn write_add_in(extension: &AddIn, conformance: Conformance) -> Result<Vec<u8>> {
    write_add_in_with(extension, conformance, &Limits::standard())
}

pub(super) fn write_add_in_with(
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
pub(super) fn write_panes(task_panes: &Panes, conformance: Conformance) -> Result<Vec<u8>> {
    write_panes_with(task_panes, conformance, &Limits::standard())
}

pub(super) fn write_panes_with(
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

pub(super) fn parse_task_pane(node: &Node, document: &XmlDocument) -> Result<ParsedPane> {
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

pub(super) fn parse_store_reference(node: &Node, document: &XmlDocument) -> Result<Reference> {
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

pub(super) fn parse_property(node: &Node) -> Result<Property> {
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

pub(super) fn parse_binding(node: &Node, document: &XmlDocument) -> Result<Binding> {
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

pub(super) fn load_snapshot_resources(
    package: &OpcPackage,
    part: &dyn Part,
    extension: &AddIn,
    total_snapshot_bytes: &mut usize,
    counted_snapshot_parts: &mut HashSet<String>,
    limits: &Limits,
    index: &PackageGraphIndex,
) -> Result<Vec<SnapshotResource>> {
    let mut referenced = HashMap::new();
    if let Some(snapshot) = &extension.snapshot {
        if let Some(id) = &snapshot.embedded_relationship_id
            && referenced.insert(id.as_str(), false).is_some()
        {
            return invalid("snapshot embed and link IDs must differ".into());
        }
        if let Some(id) = &snapshot.linked_relationship_id
            && referenced.insert(id.as_str(), true).is_some()
        {
            return invalid("snapshot embed and link IDs must differ".into());
        }
    }
    let mut resources = Vec::with_capacity(referenced.len());
    for relationship in part.rels().iter() {
        let Some(linked) = referenced.remove(relationship.r_id()) else {
            return invalid(format!(
                "web extension part has unreferenced relationship '{}'",
                relationship.r_id()
            ));
        };
        if !matches!(
            relationship.reltype(),
            IMAGE_RELATIONSHIP_TYPE | STRICT_IMAGE_RELATIONSHIP_TYPE
        ) {
            return invalid(format!(
                "snapshot relationship '{}' is not an image relationship",
                relationship.r_id()
            ));
        }
        if relationship.is_external() {
            if !linked {
                return invalid(format!(
                    "embedded snapshot relationship '{}' must be internal",
                    relationship.r_id()
                ));
            }
            resources.push(SnapshotResource {
                relationship_id: relationship.r_id().to_owned(),
                target: SnapshotTarget::External {
                    target: relationship.target_ref().to_owned(),
                },
            });
            continue;
        }
        let image_target = checked_internal_target(relationship, "snapshot image")?;
        let image_name = index
            .canonical(&image_target)
            .ok_or_else(|| Error::Missing(format!("snapshot image '{}'", image_target.as_str())))?;
        let image = package.get_part(image_name).map_err(|error| {
            Error::Missing(format!("snapshot image '{}': {error}", image_name.as_str()))
        })?;
        validate_image_content_type(image.content_type())?;
        if image.rels().iter().next().is_some() {
            return invalid(format!(
                "snapshot image '{}' must not have relationships",
                image_name.as_str()
            ));
        }
        if image.blob().len() > limits.image_bytes {
            return limit(
                "web extension snapshot bytes",
                limits.image_bytes,
                image.blob().len(),
            );
        }
        let image_name = image.partname().clone();
        if counted_snapshot_parts.insert(fold_part_name(&image_name)) {
            *total_snapshot_bytes = total_snapshot_bytes
                .checked_add(image.blob().len())
                .ok_or_else(|| Error::Invalid("aggregate snapshot byte count overflow".into()))?;
            if *total_snapshot_bytes > limits.total_image_bytes {
                return limit(
                    "aggregate web extension snapshot bytes",
                    limits.total_image_bytes,
                    *total_snapshot_bytes,
                );
            }
        }
        resources.push(SnapshotResource {
            relationship_id: relationship.r_id().to_owned(),
            target: SnapshotTarget::Internal {
                part_name: image_name,
                content_type: image.content_type().to_owned(),
                data: image.blob_arc(),
            },
        });
    }
    if let Some((id, _)) = referenced.into_iter().next() {
        return invalid(format!("snapshot references missing relationship '{id}'"));
    }
    let embedded_id = extension
        .snapshot
        .as_ref()
        .and_then(|snapshot| snapshot.embedded_relationship_id.as_deref());
    resources.sort_by(|left, right| {
        let left_order = usize::from(Some(left.relationship_id.as_str()) != embedded_id);
        let right_order = usize::from(Some(right.relationship_id.as_str()) != embedded_id);
        left_order
            .cmp(&right_order)
            .then_with(|| left.relationship_id.cmp(&right.relationship_id))
    });
    Ok(resources)
}

pub(super) fn validate_model(extension: &AddIn) -> Result<()> {
    validate_model_with(extension, &Limits::standard())
}

pub(super) fn validate_model_with(extension: &AddIn, limits: &Limits) -> Result<()> {
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

pub(super) fn validate_binding(binding: &Binding) -> Result<()> {
    require_nonempty("binding id", &binding.id)?;
    require_nonempty("binding type", binding.kind.as_str())?;
    require_nonempty("binding appref", &binding.app_ref)?;
    validate_extension_list(binding.extension_list.as_ref(), &[ExtKind::AddIn])
}

pub(super) fn validate_store_reference(reference: &Reference) -> Result<()> {
    require_nonempty("reference id", &reference.id)?;
    require_nonempty("reference version", &reference.version)?;
    if let Some(location) = &reference.location {
        require_nonempty("reference location", location)?;
    } else if reference.store == Store::FileSystem {
        return invalid("FileSystem reference requires a non-empty location".into());
    }
    validate_extension_list(reference.extension_list.as_ref(), &[ExtKind::AddIn])
}

pub(super) fn validate_task_pane(pane: &Pane) -> Result<()> {
    validate_task_pane_with(pane, &Limits::standard())
}

pub(super) fn validate_task_pane_with(pane: &Pane, limits: &Limits) -> Result<()> {
    require_nonempty("dock state", pane.dock_state.as_str())?;
    require_nonempty("task-pane relationship id", &pane.relationship_id)?;
    if !pane.width.is_finite() || pane.width <= 0.0 {
        return invalid("task-pane width must be finite and positive".into());
    }
    validate_extension_list(pane.extension_list.as_ref(), &[ExtKind::TaskPane])?;
    validate_model_with(&pane.add_in, limits)?;
    validate_snapshot_resources_with(pane, limits)
}

pub(super) fn validate_panes(task_panes: &Panes, limits: &Limits) -> Result<()> {
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

pub(super) fn charge_authored_metadata(
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

pub(super) fn add_xml_budget(total: &mut usize, bytes: usize, limits: &Limits) -> Result<()> {
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

pub(super) fn add_escaped_xml_budget(
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

pub(super) fn add_reference_budget(
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

pub(super) fn validate_add_in_budget(extension: &AddIn, limits: &Limits) -> Result<()> {
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

pub(super) fn validate_panes_budget(task_panes: &Panes, limits: &Limits) -> Result<()> {
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

pub(super) fn validate_extension_list(
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

pub(super) fn validate_snapshot_resources_with(pane: &Pane, limits: &Limits) -> Result<()> {
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

pub(super) fn write_store_reference(out: &mut String, element: &str, reference: &Reference) {
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

pub(super) fn format_f64(value: f64) -> String {
    let mut buffer = ryu::Buffer::new();
    buffer.format_finite(value).to_owned()
}

pub(super) fn escape_attr(out: &mut String, value: &str) {
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

pub(super) fn effective_namespaces(scope: &NamespaceScope) -> Result<Vec<(&String, &String)>> {
    let mut namespaces = Vec::new();
    namespaces
        .try_reserve(scope.binding_count)
        .map_err(|_| Error::Limit {
            resource: "retained web extension namespace entries",
            max: scope.binding_count,
            actual: scope.binding_count,
        })?;
    let mut seen = HashSet::new();
    seen.try_reserve(scope.binding_count)
        .map_err(|_| Error::Limit {
            resource: "retained web extension namespace entries",
            max: scope.binding_count,
            actual: scope.binding_count,
        })?;
    let mut current = Some(scope);
    while let Some(value) = current {
        for (prefix, namespace) in &value.local {
            if seen.insert(prefix.as_str()) {
                namespaces.push((prefix, namespace));
            }
        }
        current = value.parent.as_deref();
    }
    Ok(namespaces)
}

pub(super) fn retained_namespace_bytes(
    namespaces: &[(&String, &String)],
    declared_prefixes: &HashSet<String>,
) -> Result<usize> {
    let mut total = 0usize;
    for (prefix, namespace) in namespaces {
        if prefix.as_str() == "xml" || declared_prefixes.contains(prefix.as_str()) {
            continue;
        }
        let head = if prefix.is_empty() {
            " xmlns=\"".len()
        } else {
            " xmlns:"
                .len()
                .checked_add(prefix.len())
                .and_then(|value| value.checked_add("=\"".len()))
                .ok_or(Error::Limit {
                    resource: "retained web extension namespace bytes",
                    max: usize::MAX,
                    actual: usize::MAX,
                })?
        };
        let value = escaped_attr_bytes(namespace)?;
        total = total
            .checked_add(head)
            .and_then(|total| total.checked_add(value))
            .and_then(|total| total.checked_add(1))
            .ok_or(Error::Limit {
                resource: "retained web extension namespace bytes",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
    }
    Ok(total)
}

pub(super) fn escaped_attr_bytes(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |total, character| {
        let bytes = match character {
            '&' => "&amp;".len(),
            '<' => "&lt;".len(),
            '>' => "&gt;".len(),
            '"' => "&quot;".len(),
            '\'' => "&apos;".len(),
            _ => character.len_utf8(),
        };
        total.checked_add(bytes).ok_or(Error::Limit {
            resource: "retained web extension namespace bytes",
            max: usize::MAX,
            actual: usize::MAX,
        })
    })
}

pub(super) fn canonical_node_xml(node: &Node) -> String {
    pub(super) fn write_node(out: &mut String, node: &Node) {
        out.push('<');
        out.push_str(&node.local_name);
        out.push_str(" xmlns=\"");
        escape_attr(out, &node.namespace);
        out.push('"');
        for (index, attribute) in node.attributes.iter().enumerate() {
            if attribute.namespace.is_empty() {
                out.push(' ');
                out.push_str(&attribute.local_name);
            } else if attribute.namespace == "http://www.w3.org/XML/1998/namespace" {
                out.push_str(" xml:");
                out.push_str(&attribute.local_name);
            } else {
                out.push_str(" xmlns:n");
                out.push_str(&index.to_string());
                out.push_str("=\"");
                escape_attr(out, &attribute.namespace);
                out.push_str("\" n");
                out.push_str(&index.to_string());
                out.push(':');
                out.push_str(&attribute.local_name);
            }
            out.push_str("=\"");
            escape_attr(out, &attribute.value);
            out.push('"');
        }
        if node.children.is_empty() {
            out.push_str("/>");
            return;
        }
        out.push('>');
        for child in &node.children {
            write_node(out, child);
        }
        out.push_str("</");
        out.push_str(&node.local_name);
        out.push('>');
    }

    let mut out = String::new();
    write_node(&mut out, node);
    out
}

pub(super) fn require_nonempty(label: &str, value: &str) -> Result<()> {
    if value.is_empty() {
        invalid(format!("{label} must not be empty"))
    } else {
        Ok(())
    }
}

pub(super) fn validate_image_content_type(value: &str) -> Result<()> {
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

pub(super) fn is_mime_token_byte(byte: u8) -> bool {
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

pub(super) fn validate_external_uri_reference(value: &str) -> Result<()> {
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

pub(super) fn require_content_type(part: &dyn Part, expected: &str) -> Result<()> {
    if part.content_type() != expected {
        Err(Error::ContentType {
            expected: expected.into(),
            actual: part.content_type().into(),
        })
    } else {
        Ok(())
    }
}

pub(super) fn checked_internal_target(
    relationship: &litchi_opc::Relationship,
    label: &str,
) -> Result<PackURI> {
    if relationship.is_external() {
        return Err(Error::Relationship(format!(
            "{label} relationship '{}' must be internal",
            relationship.r_id()
        )));
    }
    if relationship.target_ref().contains(['?', '#']) {
        return Err(Error::Relationship(format!(
            "{label} relationship '{}' has an internal target with a query or fragment",
            relationship.r_id()
        )));
    }
    relationship.target_partname().map_err(|error| {
        Error::Relationship(format!(
            "invalid {label} relationship target '{}': {error}",
            relationship.r_id()
        ))
    })
}

pub(super) fn enforce_count_with(label: &'static str, count: usize, limits: &Limits) -> Result<()> {
    if count > limits.items {
        limit(label, limits.items, count)
    } else {
        Ok(())
    }
}

pub(super) fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("invalid XML boolean '{value}'")),
    }
}

pub(super) fn invalid<T>(message: String) -> Result<T> {
    Err(Error::Invalid(message))
}

pub(super) fn limit<T>(resource: &'static str, max: usize, actual: usize) -> Result<T> {
    Err(Error::Limit {
        resource,
        max,
        actual,
    })
}

#[derive(Debug)]
pub(super) struct Attribute {
    pub(super) namespace: String,
    pub(super) local_name: String,
    pub(super) value: String,
}

#[derive(Debug)]
pub(super) struct Node {
    pub(super) namespace: String,
    pub(super) local_name: String,
    pub(super) attributes: Vec<Attribute>,
    pub(super) children: Vec<Node>,
    pub(super) raw_fragment: Option<RawFragment>,
}

#[derive(Debug)]
pub(super) struct RawFragment {
    pub(super) start: usize,
    pub(super) start_tag_end: usize,
    pub(super) end: usize,
    pub(super) namespaces: Arc<NamespaceScope>,
    pub(super) declared_prefixes: HashSet<String>,
}

#[derive(Debug)]
pub(super) struct NamespaceScope {
    pub(super) parent: Option<Arc<NamespaceScope>>,
    pub(super) local: HashMap<String, String>,
    pub(super) binding_count: usize,
}

impl NamespaceScope {
    pub(super) fn xml() -> Arc<Self> {
        Arc::new(Self {
            parent: None,
            local: HashMap::from([("xml".into(), "http://www.w3.org/XML/1998/namespace".into())]),
            binding_count: 1,
        })
    }

    pub(super) fn get(&self, prefix: &str) -> Option<&str> {
        self.local
            .get(prefix)
            .map(String::as_str)
            .or_else(|| self.parent.as_deref().and_then(|parent| parent.get(prefix)))
    }
}

#[derive(Debug)]
pub(super) struct NodeFrame {
    pub(super) node: Node,
    pub(super) namespaces: Arc<NamespaceScope>,
    pub(super) extension_depth: Option<usize>,
    pub(super) direct_extension_count: usize,
}

#[derive(Debug, Default)]
pub(super) struct XmlBuildState {
    pub(super) root: Option<Node>,
    pub(super) stack: Vec<NodeFrame>,
    pub(super) string_bytes: usize,
    pub(super) nodes: usize,
}

#[derive(Debug)]
pub(super) struct XmlDocument {
    pub(super) root: Option<Node>,
    pub(super) xml: Vec<u8>,
    pub(super) string_bytes: usize,
}

impl XmlDocument {
    pub(super) fn root(&self) -> Result<&Node> {
        self.root
            .as_ref()
            .ok_or_else(|| Error::Invalid("missing XML root".into()))
    }

    pub(super) fn self_contained_fragment(&self, node: &Node) -> Result<String> {
        let fragment = node
            .raw_fragment
            .as_ref()
            .ok_or_else(|| Error::Invalid("XML node has no retained fragment bounds".into()))?;
        if fragment.start > fragment.start_tag_end
            || fragment.start_tag_end > fragment.end
            || fragment.end > self.xml.len()
        {
            return invalid("invalid retained XML fragment bounds".into());
        }
        let raw = &self.xml[fragment.start..fragment.end];
        let start_tag_end = fragment.start_tag_end - fragment.start;
        if start_tag_end == 0 || raw.get(start_tag_end - 1) != Some(&b'>') {
            return invalid("retained XML fragment has an invalid start tag".into());
        }
        let mut insert_at = start_tag_end - 1;
        let mut cursor = insert_at;
        while cursor > 0 && raw[cursor - 1].is_ascii_whitespace() {
            cursor -= 1;
        }
        if cursor > 0 && raw[cursor - 1] == b'/' {
            insert_at = cursor - 1;
        }

        let raw = std::str::from_utf8(raw)
            .map_err(|error| Error::Xml(format!("non-UTF-8 extension fragment: {error}")))?;
        let mut namespaces = effective_namespaces(&fragment.namespaces)?;
        namespaces.sort_unstable_by(|left, right| left.0.cmp(right.0));
        let extra = retained_namespace_bytes(&namespaces, &fragment.declared_prefixes)?;
        let capacity = raw.len().checked_add(extra).ok_or(Error::Limit {
            resource: "retained web extension fragment bytes",
            max: usize::MAX,
            actual: usize::MAX,
        })?;
        let mut out = String::new();
        out.try_reserve(capacity).map_err(|_| Error::Limit {
            resource: "retained web extension fragment bytes",
            max: capacity,
            actual: capacity,
        })?;
        out.push_str(&raw[..insert_at]);
        for (prefix, namespace) in namespaces {
            if prefix == "xml" || fragment.declared_prefixes.contains(prefix) {
                continue;
            }
            if prefix.is_empty() {
                out.push_str(" xmlns=\"");
            } else {
                out.push_str(" xmlns:");
                out.push_str(prefix);
                out.push_str("=\"");
            }
            escape_attr(&mut out, namespace);
            out.push('"');
        }
        out.push_str(&raw[insert_at..]);
        Ok(out)
    }
}

pub(super) fn parse_mce_xml(
    xml: &[u8],
    namespaces: &[&str],
    limits: &Limits,
) -> Result<XmlDocument> {
    if xml.len() > limits.xml_bytes {
        return limit("web extension XML bytes", limits.xml_bytes, xml.len());
    }
    let mut capabilities = MceCapabilities::ooxml_baseline();
    for namespace in namespaces {
        capabilities.understand_namespace(*namespace);
    }
    let mce_limits = MceLimits {
        max_input_bytes: limits.xml_bytes,
        max_output_bytes: limits.xml_bytes,
        max_depth: limits.depth,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    };
    let processed = process_markup_compatibility(xml, &capabilities, &mce_limits)?;
    parse_xml_owned(processed.xml.into_owned(), limits)
}

pub(super) fn parse_xml(xml: &[u8]) -> Result<XmlDocument> {
    let limits = Limits::standard();
    if xml.len() > limits.xml_bytes {
        return limit("web extension XML bytes", limits.xml_bytes, xml.len());
    }
    parse_xml_owned(xml.to_vec(), &limits)
}

pub(super) fn parse_xml_owned(xml: Vec<u8>, limits: &Limits) -> Result<XmlDocument> {
    if xml.len() > limits.xml_bytes {
        return limit("web extension XML bytes", limits.xml_bytes, xml.len());
    }
    let mut reader = Reader::from_reader(xml.as_slice());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let mut state = XmlBuildState::default();
    let mut xml_version = XmlVersion::Implicit1_0;
    let mut declaration_seen = false;
    let mut content_seen = false;
    loop {
        let event_start = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buffer)?;
        let event_end = reader.buffer_position() as usize;
        let declaration_or_eof = matches!(&event, Event::Decl(_) | Event::Eof);
        match event {
            Event::Decl(declaration) => {
                if declaration_seen || content_seen {
                    return invalid("XML declaration must appear once at the beginning".into());
                }
                declaration_seen = true;
                xml_version = declaration.xml_version()?;
                if xml_version == XmlVersion::Explicit1_1 {
                    return invalid("XML 1.1 is not supported for web extension parts".into());
                }
            },
            Event::Start(element) => push_element(
                &reader,
                &element,
                &mut state,
                xml_version,
                ElementEvent {
                    empty: false,
                    start: event_start,
                    end: event_end,
                },
                limits,
            )?,
            Event::Empty(element) => push_element(
                &reader,
                &element,
                &mut state,
                xml_version,
                ElementEvent {
                    empty: true,
                    start: event_start,
                    end: event_end,
                },
                limits,
            )?,
            Event::Eof => break,
            Event::DocType(_) => return invalid("DTD is forbidden in web extension XML".into()),
            Event::Text(text)
                if !extension_text_is_allowed(&state.stack)
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return invalid("text is not permitted in web extension structures".into());
            },
            Event::CData(text)
                if !extension_text_is_allowed(&state.stack)
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return invalid("CDATA is not permitted in web extension structures".into());
            },
            Event::GeneralRef(_) => {
                return invalid(
                    "general entity references are forbidden in web extension XML".into(),
                );
            },
            Event::End(_) if state.stack.is_empty() => {
                return invalid("unexpected XML end tag".into());
            },
            Event::End(_) => {
                let mut frame = state
                    .stack
                    .pop()
                    .ok_or_else(|| Error::Invalid("unexpected XML end tag".into()))?;
                if let Some(fragment) = frame.node.raw_fragment.as_mut() {
                    fragment.end = event_end;
                }
                attach_node(&mut state.root, &mut state.stack, frame.node)?;
            },
            _ => {},
        }
        if !declaration_or_eof {
            content_seen = true;
        }
        buffer.clear();
    }
    if !state.stack.is_empty() {
        return invalid("unclosed XML element".into());
    }
    if state.string_bytes > limits.string_bytes {
        return limit(
            "web extension decoded string bytes",
            limits.string_bytes,
            state.string_bytes,
        );
    }
    drop(reader);
    Ok(XmlDocument {
        root: state.root,
        xml,
        string_bytes: state.string_bytes,
    })
}

#[derive(Debug, Clone, Copy)]
pub(super) struct ElementEvent {
    pub(super) empty: bool,
    pub(super) start: usize,
    pub(super) end: usize,
}

pub(super) fn push_element(
    reader: &Reader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    state: &mut XmlBuildState,
    xml_version: XmlVersion,
    event: ElementEvent,
    limits: &Limits,
) -> Result<()> {
    if state.stack.len() >= limits.depth {
        return limit(
            "web extension XML depth",
            limits.depth,
            state.stack.len().saturating_add(1),
        );
    }
    state.nodes = state
        .nodes
        .checked_add(1)
        .ok_or_else(|| Error::Invalid("web extension node count overflow".into()))?;
    if state.nodes > limits.nodes {
        return limit("web extension XML nodes", limits.nodes, state.nodes);
    }
    let parent_namespaces = state
        .stack
        .last()
        .map(|frame| Arc::clone(&frame.namespaces))
        .unwrap_or_else(NamespaceScope::xml);
    let mut local_namespaces = HashMap::new();
    let mut raw_attributes = Vec::new();
    let mut declared_prefixes = HashSet::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(xml_version, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        state.string_bytes = state
            .string_bytes
            .checked_add(name.len().saturating_add(value.len()))
            .ok_or(Error::Limit {
                resource: "web extension decoded string bytes",
                max: limits.string_bytes,
                actual: usize::MAX,
            })?;
        if state.string_bytes > limits.string_bytes {
            return limit(
                "web extension decoded string bytes",
                limits.string_bytes,
                state.string_bytes,
            );
        }
        if name == "xmlns" {
            if !declared_prefixes.insert(String::new()) {
                return invalid("duplicate default namespace declaration".into());
            }
            local_namespaces.insert(String::new(), value);
        } else if let Some(prefix) = name.strip_prefix("xmlns:") {
            if prefix == "xmlns"
                || (prefix == "xml" && value != "http://www.w3.org/XML/1998/namespace")
                || value.is_empty()
            {
                return invalid(format!(
                    "invalid namespace declaration for prefix '{prefix}'"
                ));
            }
            if !declared_prefixes.insert(prefix.to_owned()) {
                return invalid(format!(
                    "duplicate namespace declaration for prefix '{prefix}'"
                ));
            }
            local_namespaces.insert(prefix.to_owned(), value);
        } else {
            raw_attributes.push((name, value));
        }
    }
    let new_bindings = local_namespaces
        .keys()
        .filter(|prefix| parent_namespaces.get(prefix).is_none())
        .count();
    let binding_count = parent_namespaces
        .binding_count
        .checked_add(new_bindings)
        .ok_or(Error::Limit {
            resource: "web extension XML namespace bindings",
            max: 4096,
            actual: usize::MAX,
        })?;
    if binding_count > 4096 {
        return invalid("web extension XML namespace bindings exceed 4096".into());
    }
    let namespaces = if local_namespaces.is_empty() {
        parent_namespaces
    } else {
        Arc::new(NamespaceScope {
            parent: Some(parent_namespaces),
            local: local_namespaces,
            binding_count,
        })
    };
    let element_name = element.name();
    let raw_name = std::str::from_utf8(element_name.as_ref())
        .map_err(|error| Error::Xml(error.to_string()))?;
    let (prefix, local_name) = split_qname(raw_name);
    let namespace = if prefix.is_empty() {
        namespaces.get(prefix).unwrap_or_default().to_owned()
    } else {
        namespaces
            .get(prefix)
            .map(str::to_owned)
            .ok_or_else(|| Error::Invalid(format!("unbound XML namespace prefix '{prefix}'")))?
    };
    state.string_bytes = state
        .string_bytes
        .checked_add(namespace.len().saturating_add(local_name.len()))
        .ok_or(Error::Limit {
            resource: "web extension decoded string bytes",
            max: limits.string_bytes,
            actual: usize::MAX,
        })?;
    if state.string_bytes > limits.string_bytes {
        return limit(
            "web extension decoded string bytes",
            limits.string_bytes,
            state.string_bytes,
        );
    }
    let mut attributes = Vec::with_capacity(raw_attributes.len());
    let mut seen = HashSet::new();
    for (raw_name, value) in raw_attributes {
        let (prefix, local_name) = split_qname(&raw_name);
        let namespace = if prefix.is_empty() {
            String::new()
        } else {
            namespaces
                .get(prefix)
                .map(str::to_owned)
                .ok_or_else(|| Error::Invalid(format!("unbound attribute prefix '{prefix}'")))?
        };
        if !seen.insert((namespace.clone(), local_name.to_owned())) {
            return invalid(format!("duplicate attribute {{{namespace}}}{local_name}"));
        }
        attributes.push(Attribute {
            namespace,
            local_name: local_name.to_owned(),
            value,
        });
    }
    let capture_fragment = should_capture_extension_list(
        state.stack.last().map(|frame| &frame.node),
        &namespace,
        local_name,
    );
    let raw_fragment = if capture_fragment {
        let inherited = effective_namespaces(&namespaces)?;
        let retained_bytes = declared_prefixes.iter().try_fold(
            retained_namespace_bytes(&inherited, &declared_prefixes)?,
            |total, prefix| {
                total.checked_add(prefix.len()).ok_or(Error::Limit {
                    resource: "web extension decoded string bytes",
                    max: limits.string_bytes,
                    actual: usize::MAX,
                })
            },
        )?;
        state.string_bytes =
            state
                .string_bytes
                .checked_add(retained_bytes)
                .ok_or(Error::Limit {
                    resource: "web extension decoded string bytes",
                    max: limits.string_bytes,
                    actual: usize::MAX,
                })?;
        if state.string_bytes > limits.string_bytes {
            return limit(
                "web extension decoded string bytes",
                limits.string_bytes,
                state.string_bytes,
            );
        }
        Some(RawFragment {
            start: event.start,
            start_tag_end: event.end,
            end: if event.empty { event.end } else { 0 },
            namespaces: Arc::clone(&namespaces),
            declared_prefixes,
        })
    } else {
        None
    };
    let node = Node {
        namespace,
        local_name: local_name.to_owned(),
        attributes,
        children: Vec::new(),
        raw_fragment,
    };
    if state
        .stack
        .last()
        .is_some_and(|frame| frame.extension_depth == Some(0))
    {
        let parent = state
            .stack
            .last_mut()
            .ok_or_else(|| Error::Invalid("extension-list child has no parent element".into()))?;
        let expected_namespace = if parent.node.namespace == STRICT_DRAWINGML_NAMESPACE {
            STRICT_DRAWINGML_NAMESPACE
        } else {
            DRAWINGML_NAMESPACE
        };
        require_name(&node, expected_namespace, "ext")?;
        reject_unknown_attributes(&node, &[("", "uri")])?;
        required_attr(&node, "", "uri")?;
        parent.direct_extension_count = parent
            .direct_extension_count
            .checked_add(1)
            .ok_or_else(|| Error::Invalid("extLst count overflow".into()))?;
        enforce_count_with("OfficeArt extension", parent.direct_extension_count, limits)?;
    }
    let extension_depth = if capture_fragment {
        Some(0)
    } else {
        state
            .stack
            .last()
            .and_then(|frame| frame.extension_depth)
            .map(|depth| depth + 1)
    };
    if event.empty {
        attach_node(&mut state.root, &mut state.stack, node)?;
    } else {
        state.stack.push(NodeFrame {
            node,
            namespaces,
            extension_depth,
            direct_extension_count: 0,
        });
    }
    Ok(())
}

pub(super) fn attach_node(
    root: &mut Option<Node>,
    stack: &mut [NodeFrame],
    node: Node,
) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        if parent.extension_depth.is_none() {
            parent.node.children.push(node);
        }
    } else if root.replace(node).is_some() {
        return invalid("multiple XML root elements".into());
    }
    Ok(())
}

pub(super) fn should_capture_extension_list(
    parent: Option<&Node>,
    namespace: &str,
    local_name: &str,
) -> bool {
    if local_name != "extLst" {
        return false;
    }
    let allowed_namespace = matches!(
        namespace,
        WEB_EXTENSION_NAMESPACE
            | TASK_PANES_NAMESPACE
            | DRAWINGML_NAMESPACE
            | STRICT_DRAWINGML_NAMESPACE
    );
    if !allowed_namespace {
        return false;
    }
    let Some(parent) = parent else {
        return true;
    };
    matches!(
        (
            parent.namespace.as_str(),
            parent.local_name.as_str(),
            namespace
        ),
        (
            WEB_EXTENSION_NAMESPACE,
            "webextension" | "reference" | "binding",
            WEB_EXTENSION_NAMESPACE
        ) | (
            WEB_EXTENSION_NAMESPACE,
            "snapshot",
            DRAWINGML_NAMESPACE | STRICT_DRAWINGML_NAMESPACE
        ) | (TASK_PANES_NAMESPACE, "taskpane", TASK_PANES_NAMESPACE)
    )
}

pub(super) fn extension_text_is_allowed(stack: &[NodeFrame]) -> bool {
    stack
        .last()
        .and_then(|frame| frame.extension_depth)
        .is_some_and(|depth| depth >= 2)
}

pub(super) fn split_qname(name: &str) -> (&str, &str) {
    name.split_once(':').unwrap_or(("", name))
}

pub(super) fn element_children(node: &Node) -> Vec<&Node> {
    node.children.iter().collect()
}

pub(super) fn require_name(node: &Node, namespace: &str, local_name: &str) -> Result<()> {
    if node.namespace == namespace && node.local_name == local_name {
        Ok(())
    } else {
        invalid(format!(
            "expected {{{namespace}}}{local_name}, got {{{}}}{}",
            node.namespace, node.local_name
        ))
    }
}

pub(super) fn attr<'a>(node: &'a Node, namespace: &str, local_name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace == namespace && attribute.local_name == local_name)
        .map(|attribute| attribute.value.as_str())
}

pub(super) fn required_attr<'a>(
    node: &'a Node,
    namespace: &str,
    local_name: &str,
) -> Result<&'a str> {
    attr(node, namespace, local_name).ok_or_else(|| {
        Error::Invalid(format!(
            "{} requires attribute {{{namespace}}}{local_name}",
            node.local_name
        ))
    })
}

pub(super) fn relationship_attr<'a>(node: &'a Node, local_name: &str) -> Result<Option<&'a str>> {
    let transitional = attr(node, TRANSITIONAL_RELATIONSHIPS_NAMESPACE, local_name);
    let strict = attr(node, STRICT_RELATIONSHIPS_NAMESPACE, local_name);
    if transitional.is_some() && strict.is_some() {
        invalid(format!(
            "{} has both Strict and Transitional r:{local_name}",
            node.local_name
        ))
    } else {
        Ok(transitional.or(strict))
    }
}

pub(super) fn is_drawingml_namespace(namespace: &str) -> bool {
    matches!(namespace, DRAWINGML_NAMESPACE | STRICT_DRAWINGML_NAMESPACE)
}

pub(super) fn optional_bool_attr(
    node: &Node,
    namespace: &str,
    local_name: &str,
) -> Result<Option<bool>> {
    attr(node, namespace, local_name)
        .map(parse_bool)
        .transpose()
}

pub(super) fn reject_unknown_attributes(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    for attribute in &node.attributes {
        if !allowed.iter().any(|(namespace, local_name)| {
            attribute.namespace == *namespace && attribute.local_name == *local_name
        }) {
            return invalid(format!(
                "unexpected attribute {{{}}}{} on {}",
                attribute.namespace, attribute.local_name, node.local_name
            ));
        }
    }
    Ok(())
}

pub(super) fn is_next(
    children: &[&Node],
    position: usize,
    namespace: &str,
    local_name: &str,
) -> bool {
    children
        .get(position)
        .is_some_and(|child| child.namespace == namespace && child.local_name == local_name)
}

pub(super) fn next_required<'a>(
    children: &[&'a Node],
    position: &mut usize,
    namespace: &str,
    local_name: &str,
) -> Result<&'a Node> {
    if !is_next(children, *position, namespace, local_name) {
        return invalid(format!("missing or misplaced {local_name}"));
    }
    let node = children[*position];
    *position += 1;
    Ok(node)
}

pub(super) fn ensure_consumed(children: &[&Node], position: usize, parent: &str) -> Result<()> {
    if position == children.len() {
        Ok(())
    } else {
        invalid(format!(
            "unexpected child {} in {parent}",
            children[position].local_name
        ))
    }
}
