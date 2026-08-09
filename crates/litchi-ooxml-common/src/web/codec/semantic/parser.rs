use super::super::super::model::{
    AddIn, Binding, BindingKind, Compression, Dock, Effect, ExtList, Limits, OperationBudget,
    Property, Reference, Snapshot, Store,
};
use super::super::super::validation::validate_store_reference;
use super::super::super::{
    STRICT_RELATIONSHIPS_NAMESPACE, TASK_PANES_NAMESPACE, TRANSITIONAL_RELATIONSHIPS_NAMESPACE,
    WEB_EXTENSION_NAMESPACE,
};
use super::super::relationship::relationship_attr;
use super::super::xml::{
    Node, XmlDocument, attr, element_children, ensure_consumed, is_drawingml_namespace, is_next,
    next_required, optional_bool_attr, parse_mce_xml, reject_unknown_attributes, require_name,
    required_attr,
};
use super::support::{enforce_count_with, invalid, parse_bool};
use crate::{Error, Result};

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
        .map_err(|error| Error::Invalid(format!("invalid task-pane width: {error}")))?;
    if !width.is_finite() {
        return invalid("task-pane width must be finite".into());
    }
    let row = required_attr(node, "", "row")?
        .parse::<u32>()
        .map_err(|error| Error::Invalid(format!("invalid task-pane row: {error}")))?;
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
