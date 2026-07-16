//! Inert Office Web Extension and persisted task-pane metadata.
//!
//! This module implements the package structures defined by MS-OWEXML. It
//! intentionally does not locate add-ins, contact catalog providers, load
//! manifests, resolve linked content, or execute scripts/custom functions.

use crate::common::{MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};
use litchi_opc::{OpcPackage, Part};
use quick_xml::Reader;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use std::collections::{HashMap, HashSet};

pub const WEB_EXTENSION_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/webextensions/webextension/2010/11";
pub const TASK_PANES_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11";
pub const TRANSITIONAL_RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub const STRICT_RELATIONSHIPS_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships";

pub const TASK_PANES_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes";
pub const WEB_EXTENSION_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/webextension";
pub const TASK_PANES_CONTENT_TYPE: &str =
    "application/vnd.ms-office.webextensiontaskpanes+xml";
pub const WEB_EXTENSION_CONTENT_TYPE: &str =
    "application/vnd.ms-office.webextension+xml";

const IMAGE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const STRICT_IMAGE_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image";

/// Maximum bytes accepted for either XML part before MCE processing.
pub const MAX_WEB_EXTENSION_XML_BYTES: usize = 4 * 1024 * 1024;
/// Maximum XML nesting depth.
pub const MAX_WEB_EXTENSION_XML_DEPTH: usize = 128;
/// Maximum task panes, alternate references, properties, or bindings per list.
pub const MAX_WEB_EXTENSION_ITEMS: usize = 4096;
/// Maximum aggregate decoded string bytes in one XML part.
pub const MAX_WEB_EXTENSION_STRING_BYTES: usize = 8 * 1024 * 1024;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OoxmlConformance {
    Transitional,
    Strict,
}

impl OoxmlConformance {
    fn relationships_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_RELATIONSHIPS_NAMESPACE,
            Self::Strict => STRICT_RELATIONSHIPS_NAMESPACE,
        }
    }
}

/// Catalog provider type from MS-OWEXML section 2.2.5.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum WebExtensionStoreType {
    Omex,
    #[default]
    SharePointCatalog,
    SharePointApp,
    Exchange,
    FileSystem,
    Registry,
    ExchangeCatalog,
    WopiCatalog,
}

impl WebExtensionStoreType {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Omex => "OMEX",
            Self::SharePointCatalog => "SPCatalog",
            Self::SharePointApp => "SPApp",
            Self::Exchange => "Exchange",
            Self::FileSystem => "FileSystem",
            Self::Registry => "Registry",
            Self::ExchangeCatalog => "ExCatalog",
            Self::WopiCatalog => "WOPICatalog",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "OMEX" => Ok(Self::Omex),
            "SPCatalog" => Ok(Self::SharePointCatalog),
            "SPApp" => Ok(Self::SharePointApp),
            "Exchange" => Ok(Self::Exchange),
            "FileSystem" => Ok(Self::FileSystem),
            "Registry" => Ok(Self::Registry),
            "ExCatalog" => Ok(Self::ExchangeCatalog),
            "WOPICatalog" => Ok(Self::WopiCatalog),
            _ => invalid(format!("invalid web extension storeType '{value}'")),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WebExtensionStoreReference {
    pub id: String,
    pub version: String,
    pub store: Option<String>,
    pub store_type: WebExtensionStoreType,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WebExtensionProperty {
    pub name: String,
    pub value: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WebExtensionBinding {
    pub id: String,
    pub binding_type: String,
    pub application_reference: String,
}

/// Relationship-bearing subset of DrawingML `CT_Blip` used by a snapshot.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct WebExtensionSnapshot {
    pub embedded_relationship_id: Option<String>,
    pub linked_relationship_id: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WebExtension {
    pub id: String,
    pub frozen: bool,
    pub reference: WebExtensionStoreReference,
    pub alternate_references: Vec<WebExtensionStoreReference>,
    pub properties: Vec<WebExtensionProperty>,
    pub bindings: Vec<WebExtensionBinding>,
    pub snapshot: Option<WebExtensionSnapshot>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct WebExtensionTaskPane {
    /// Schema-defined free-form docking state, commonly `left` or `right`.
    pub dock_state: String,
    pub visible: bool,
    pub width: f64,
    pub row: u32,
    pub locked: bool,
    pub relationship_id: String,
    pub web_extension: WebExtension,
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct WebExtensionTaskPanes {
    pub panes: Vec<WebExtensionTaskPane>,
}

/// Parse one MS-OWEXML web extension part after bounded MCE preprocessing.
pub fn parse_web_extension(xml: &[u8]) -> Result<WebExtension> {
    let document = parse_mce_xml(xml, &[WEB_EXTENSION_NAMESPACE])?;
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
    let reference = parse_store_reference(reference_node)?;

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
        enforce_count("alternate reference", refs.len())?;
        refs.into_iter()
            .map(|child| {
                require_name(child, WEB_EXTENSION_NAMESPACE, "reference")?;
                parse_store_reference(child)
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
    enforce_count("property", property_nodes.len())?;
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
    enforce_count("binding", binding_nodes.len())?;
    let bindings = binding_nodes
        .into_iter()
        .map(parse_binding)
        .collect::<Result<Vec<_>>>()?;

    let snapshot = if is_next(&children, position, WEB_EXTENSION_NAMESPACE, "snapshot") {
        let node = children[position];
        position += 1;
        let embedded_relationship_id = relationship_attr(node, "embed")?.map(str::to_owned);
        let linked_relationship_id = relationship_attr(node, "link")?.map(str::to_owned);
        for attribute in &node.attributes {
            let known_relationship = is_relationship_namespace(&attribute.namespace)
                && matches!(attribute.local_name.as_str(), "embed" | "link");
            // CT_Blip has additional unqualified compression/state attributes.
            if !known_relationship && !attribute.namespace.is_empty() {
                return invalid(format!(
                    "unexpected namespaced snapshot attribute {{{}}}{}",
                    attribute.namespace, attribute.local_name
                ));
            }
        }
        Some(WebExtensionSnapshot {
            embedded_relationship_id,
            linked_relationship_id,
        })
    } else {
        None
    };

    if is_next(&children, position, WEB_EXTENSION_NAMESPACE, "extLst")
        || is_next(
            &children,
            position,
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            "extLst",
        )
        || is_next(
            &children,
            position,
            "http://purl.oclc.org/ooxml/drawingml/main",
            "extLst",
        )
    {
        position += 1;
    }
    ensure_consumed(&children, position, "webextension")?;

    Ok(WebExtension {
        id,
        frozen,
        reference,
        alternate_references,
        properties,
        bindings,
        snapshot,
    })
}

/// Parse task-pane metadata without resolving its web-extension relationships.
pub fn parse_task_panes(xml: &[u8]) -> Result<Vec<ParsedTaskPane>> {
    let document = parse_mce_xml(xml, &[TASK_PANES_NAMESPACE, WEB_EXTENSION_NAMESPACE])?;
    let root = document.root()?;
    require_name(root, TASK_PANES_NAMESPACE, "taskpanes")?;
    reject_unknown_attributes(root, &[])?;
    let children = element_children(root);
    enforce_count("task pane", children.len())?;
    children.into_iter().map(parse_task_pane).collect()
}

#[derive(Debug, Clone, PartialEq)]
pub struct ParsedTaskPane {
    pub dock_state: String,
    pub visible: bool,
    pub width: f64,
    pub row: u32,
    pub locked: bool,
    pub relationship_id: String,
}

/// Resolve and validate the complete package graph for persisted task panes.
pub fn load_web_extension_task_panes(
    package: &OpcPackage,
) -> Result<Option<WebExtensionTaskPanes>> {
    let relationships: Vec<_> = package
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP_TYPE)
        .collect();
    if relationships.is_empty() {
        return Ok(None);
    }
    if relationships.len() != 1 {
        return invalid("package has multiple web extension task-pane relationships".into());
    }
    let relationship = relationships[0];
    if relationship.is_external() {
        return invalid("task-pane relationship must be internal".into());
    }
    let part_name = relationship.target_partname().map_err(|error| {
        OoxmlError::InvalidRelationship(format!("invalid task-pane target: {error}"))
    })?;
    let task_panes_part = package.get_part(&part_name).map_err(|error| {
        OoxmlError::PartNotFound(format!("task-pane part '{}': {error}", part_name.as_str()))
    })?;
    require_content_type(task_panes_part, TASK_PANES_CONTENT_TYPE)?;
    let parsed_panes = parse_task_panes(task_panes_part.blob())?;

    let referenced_ids: HashSet<&str> = parsed_panes
        .iter()
        .map(|pane| pane.relationship_id.as_str())
        .collect();
    for child_relationship in task_panes_part.rels().iter() {
        if child_relationship.reltype() != WEB_EXTENSION_RELATIONSHIP_TYPE {
            return invalid(format!(
                "task-pane part has forbidden relationship '{}' of type '{}'",
                child_relationship.r_id(),
                child_relationship.reltype()
            ));
        }
        if !referenced_ids.contains(child_relationship.r_id()) {
            return invalid(format!(
                "task-pane part has unreferenced relationship '{}'",
                child_relationship.r_id()
            ));
        }
    }

    let mut panes = Vec::with_capacity(parsed_panes.len());
    for pane in parsed_panes {
        let child_relationship = task_panes_part
            .rels()
            .iter()
            .find(|candidate| candidate.r_id() == pane.relationship_id)
            .ok_or_else(|| {
                OoxmlError::InvalidRelationship(format!(
                    "task pane references missing relationship '{}'",
                    pane.relationship_id
                ))
            })?;
        if child_relationship.is_external() {
            return invalid(format!(
                "web extension relationship '{}' must be internal",
                pane.relationship_id
            ));
        }
        let extension_name = child_relationship.target_partname().map_err(|error| {
            OoxmlError::InvalidRelationship(format!(
                "invalid web extension target '{}': {error}",
                pane.relationship_id
            ))
        })?;
        let extension_part = package.get_part(&extension_name).map_err(|error| {
            OoxmlError::PartNotFound(format!(
                "web extension part '{}': {error}",
                extension_name.as_str()
            ))
        })?;
        require_content_type(extension_part, WEB_EXTENSION_CONTENT_TYPE)?;
        let web_extension = parse_web_extension(extension_part.blob())?;
        validate_snapshot_relationships(package, extension_part, &web_extension)?;
        panes.push(WebExtensionTaskPane {
            dock_state: pane.dock_state,
            visible: pane.visible,
            width: pane.width,
            row: pane.row,
            locked: pane.locked,
            relationship_id: pane.relationship_id,
            web_extension,
        });
    }
    Ok(Some(WebExtensionTaskPanes { panes }))
}

/// Deterministically serialize a single web extension part.
pub fn write_web_extension(
    extension: &WebExtension,
    conformance: OoxmlConformance,
) -> Result<Vec<u8>> {
    validate_model(extension)?;
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
        escape_attr(&mut out, &binding.binding_type);
        out.push_str("\" appref=\"");
        escape_attr(&mut out, &binding.application_reference);
        out.push_str("\"/>");
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
        out.push_str("/>");
    }
    out.push_str("</we:webextension>");
    Ok(out.into_bytes())
}

/// Deterministically serialize task-pane metadata and relationship IDs.
pub fn write_task_panes(
    task_panes: &WebExtensionTaskPanes,
    conformance: OoxmlConformance,
) -> Result<Vec<u8>> {
    enforce_count("task pane", task_panes.panes.len())?;
    let mut out = String::with_capacity(512 + task_panes.panes.len() * 160);
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    out.push_str("<wetp:taskpanes xmlns:wetp=\"");
    escape_attr(&mut out, TASK_PANES_NAMESPACE);
    out.push_str("\" xmlns:r=\"");
    escape_attr(&mut out, conformance.relationships_namespace());
    out.push_str("\">");
    for pane in &task_panes.panes {
        validate_task_pane(pane)?;
        out.push_str("<wetp:taskpane dockstate=\"");
        escape_attr(&mut out, &pane.dock_state);
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
        out.push_str("\"/></wetp:taskpane>");
    }
    out.push_str("</wetp:taskpanes>");
    Ok(out.into_bytes())
}

fn parse_task_pane(node: &Node) -> Result<ParsedTaskPane> {
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
    let dock_state = required_attr(node, "", "dockstate")?.to_owned();
    let visible = parse_bool(required_attr(node, "", "visibility")?)?;
    let width = required_attr(node, "", "width")?
        .parse::<f64>()
        .map_err(|_| OoxmlError::InvalidFormat("invalid task-pane width".into()))?;
    if !width.is_finite() {
        return invalid("task-pane width must be finite".into());
    }
    let row = required_attr(node, "", "row")?
        .parse::<u32>()
        .map_err(|_| OoxmlError::InvalidFormat("invalid task-pane row".into()))?;
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
        .ok_or_else(|| OoxmlError::InvalidFormat("webextensionref requires r:id".into()))?
        .to_owned();
    if children.len() > 2
        || (children.len() == 2
            && !matches!(children[1].local_name.as_str(), "extLst" | "float"))
    {
        return invalid("unexpected taskpane child or child order".into());
    }
    Ok(ParsedTaskPane {
        dock_state,
        visible,
        width,
        row,
        locked,
        relationship_id,
    })
}

fn parse_store_reference(node: &Node) -> Result<WebExtensionStoreReference> {
    require_name(node, WEB_EXTENSION_NAMESPACE, "reference")?;
    reject_unknown_attributes(
        node,
        &[("", "id"), ("", "version"), ("", "store"), ("", "storeType")],
    )?;
    let children = element_children(node);
    if children.len() > 1 || children.first().is_some_and(|child| child.local_name != "extLst") {
        return invalid("reference permits only one trailing extLst".into());
    }
    Ok(WebExtensionStoreReference {
        id: required_attr(node, "", "id")?.to_owned(),
        version: required_attr(node, "", "version")?.to_owned(),
        store: attr(node, "", "store").map(str::to_owned),
        store_type: attr(node, "", "storeType")
            .map(WebExtensionStoreType::parse)
            .transpose()?
            .unwrap_or_default(),
    })
}

fn parse_property(node: &Node) -> Result<WebExtensionProperty> {
    require_name(node, WEB_EXTENSION_NAMESPACE, "property")?;
    reject_unknown_attributes(node, &[("", "name"), ("", "value")])?;
    if !element_children(node).is_empty() {
        return invalid("web extension property must be empty".into());
    }
    Ok(WebExtensionProperty {
        name: required_attr(node, "", "name")?.to_owned(),
        value: required_attr(node, "", "value")?.to_owned(),
    })
}

fn parse_binding(node: &Node) -> Result<WebExtensionBinding> {
    require_name(node, WEB_EXTENSION_NAMESPACE, "binding")?;
    reject_unknown_attributes(node, &[("", "id"), ("", "type"), ("", "appref")])?;
    let children = element_children(node);
    if children.len() > 1 || children.first().is_some_and(|child| child.local_name != "extLst") {
        return invalid("binding permits only one trailing extLst".into());
    }
    Ok(WebExtensionBinding {
        id: required_attr(node, "", "id")?.to_owned(),
        binding_type: required_attr(node, "", "type")?.to_owned(),
        application_reference: required_attr(node, "", "appref")?.to_owned(),
    })
}

fn validate_snapshot_relationships(
    package: &OpcPackage,
    part: &dyn Part,
    extension: &WebExtension,
) -> Result<()> {
    let mut referenced = HashSet::new();
    if let Some(snapshot) = &extension.snapshot {
        if let Some(id) = &snapshot.embedded_relationship_id {
            referenced.insert(id.as_str());
        }
        if let Some(id) = &snapshot.linked_relationship_id {
            referenced.insert(id.as_str());
        }
    }
    for relationship in part.rels().iter() {
        if !referenced.contains(relationship.r_id()) {
            return invalid(format!(
                "web extension part has unreferenced relationship '{}'",
                relationship.r_id()
            ));
        }
        if !matches!(relationship.reltype(), IMAGE_RELATIONSHIP_TYPE | STRICT_IMAGE_RELATIONSHIP_TYPE)
        {
            return invalid(format!(
                "snapshot relationship '{}' is not an image relationship",
                relationship.r_id()
            ));
        }
        if relationship.is_external() {
            return invalid(format!(
                "external snapshot relationship '{}' is not loaded",
                relationship.r_id()
            ));
        }
        let image_name = relationship.target_partname().map_err(|error| {
            OoxmlError::InvalidRelationship(format!("invalid snapshot target: {error}"))
        })?;
        let image = package.get_part(&image_name).map_err(|error| {
            OoxmlError::PartNotFound(format!("snapshot image '{}': {error}", image_name.as_str()))
        })?;
        if !image.content_type().starts_with("image/") {
            return invalid(format!(
                "snapshot target '{}' has non-image content type '{}'",
                image_name.as_str(),
                image.content_type()
            ));
        }
    }
    for id in referenced {
        if !part.rels().iter().any(|relationship| relationship.r_id() == id) {
            return invalid(format!("snapshot references missing relationship '{id}'"));
        }
    }
    Ok(())
}

fn validate_model(extension: &WebExtension) -> Result<()> {
    require_nonempty("web extension id", &extension.id)?;
    validate_store_reference(&extension.reference)?;
    enforce_count("alternate reference", extension.alternate_references.len())?;
    enforce_count("property", extension.properties.len())?;
    enforce_count("binding", extension.bindings.len())?;
    for reference in &extension.alternate_references {
        validate_store_reference(reference)?;
    }
    for property in &extension.properties {
        require_nonempty("property name", &property.name)?;
    }
    for binding in &extension.bindings {
        require_nonempty("binding id", &binding.id)?;
        require_nonempty("binding type", &binding.binding_type)?;
        require_nonempty("binding appref", &binding.application_reference)?;
    }
    Ok(())
}

fn validate_store_reference(reference: &WebExtensionStoreReference) -> Result<()> {
    require_nonempty("reference id", &reference.id)?;
    require_nonempty("reference version", &reference.version)
}

fn validate_task_pane(pane: &WebExtensionTaskPane) -> Result<()> {
    require_nonempty("dock state", &pane.dock_state)?;
    require_nonempty("task-pane relationship id", &pane.relationship_id)?;
    if !pane.width.is_finite() {
        return invalid("task-pane width must be finite".into());
    }
    validate_model(&pane.web_extension)
}

fn write_store_reference(out: &mut String, element: &str, reference: &WebExtensionStoreReference) {
    out.push_str("<we:");
    out.push_str(element);
    out.push_str(" id=\"");
    escape_attr(out, &reference.id);
    out.push_str("\" version=\"");
    escape_attr(out, &reference.version);
    if let Some(store) = &reference.store {
        out.push_str("\" store=\"");
        escape_attr(out, store);
    }
    out.push_str("\" storeType=\"");
    out.push_str(reference.store_type.as_str());
    out.push_str("\"/>");
}

fn format_f64(value: f64) -> String {
    let mut buffer = ryu::Buffer::new();
    buffer.format_finite(value).to_owned()
}

fn escape_attr(out: &mut String, value: &str) {
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

fn require_nonempty(label: &str, value: &str) -> Result<()> {
    if value.is_empty() {
        invalid(format!("{label} must not be empty"))
    } else {
        Ok(())
    }
}

fn require_content_type(part: &dyn Part, expected: &str) -> Result<()> {
    if part.content_type() != expected {
        Err(OoxmlError::InvalidContentType {
            expected: expected.into(),
            got: part.content_type().into(),
        })
    } else {
        Ok(())
    }
}

fn enforce_count(label: &str, count: usize) -> Result<()> {
    if count > MAX_WEB_EXTENSION_ITEMS {
        invalid(format!(
            "{label} count {count} exceeds {MAX_WEB_EXTENSION_ITEMS}"
        ))
    } else {
        Ok(())
    }
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("invalid XML boolean '{value}'")),
    }
}

fn invalid<T>(message: String) -> Result<T> {
    Err(OoxmlError::InvalidFormat(message))
}

#[derive(Debug)]
struct Attribute {
    namespace: String,
    local_name: String,
    value: String,
}

#[derive(Debug)]
struct Node {
    namespace: String,
    local_name: String,
    attributes: Vec<Attribute>,
    children: Vec<Node>,
}

#[derive(Debug)]
struct XmlDocument {
    root: Option<Node>,
}

impl XmlDocument {
    fn root(&self) -> Result<&Node> {
        self.root
            .as_ref()
            .ok_or_else(|| OoxmlError::InvalidFormat("missing XML root".into()))
    }
}

fn parse_mce_xml(xml: &[u8], namespaces: &[&str]) -> Result<XmlDocument> {
    if xml.len() > MAX_WEB_EXTENSION_XML_BYTES {
        return invalid(format!(
            "web extension XML exceeds {MAX_WEB_EXTENSION_XML_BYTES} bytes"
        ));
    }
    let mut capabilities = MceCapabilities::ooxml_baseline();
    for namespace in namespaces {
        capabilities.understand_namespace(*namespace);
    }
    let limits = MceLimits {
        max_input_bytes: MAX_WEB_EXTENSION_XML_BYTES,
        max_output_bytes: MAX_WEB_EXTENSION_XML_BYTES * 2,
        max_depth: MAX_WEB_EXTENSION_XML_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    };
    let processed = process_markup_compatibility(xml, &capabilities, &limits)?;
    parse_xml(processed.xml.as_ref())
}

fn parse_xml(xml: &[u8]) -> Result<XmlDocument> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut document = XmlDocument { root: None };
    let mut stack: Vec<(Node, HashMap<String, String>)> = Vec::new();
    let mut string_bytes = 0usize;
    let mut xml_version = XmlVersion::Implicit1_0;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Decl(declaration) => xml_version = declaration.xml_version()?,
            Event::Start(element) => push_element(
                &reader,
                &element,
                &mut document,
                &mut stack,
                &mut string_bytes,
                xml_version,
                false,
            )?,
            Event::Empty(element) => push_element(
                &reader,
                &element,
                &mut document,
                &mut stack,
                &mut string_bytes,
                xml_version,
                true,
            )?,
            Event::Eof => break,
            Event::DocType(_) => return invalid("DTD is forbidden in web extension XML".into()),
            Event::Text(text) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                return invalid("text is not permitted in web extension structures".into());
            }
            Event::CData(text) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                return invalid("CDATA is not permitted in web extension structures".into());
            }
            Event::End(_) if stack.is_empty() => return invalid("unexpected XML end tag".into()),
            Event::End(_) => {
                let (node, _) = stack.pop().expect("stack checked above");
                attach_node(&mut document, &mut stack, node)?;
            }
            _ => {}
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return invalid("unclosed XML element".into());
    }
    if string_bytes > MAX_WEB_EXTENSION_STRING_BYTES {
        return invalid("web extension decoded strings exceed allocation limit".into());
    }
    Ok(document)
}

fn push_element(
    reader: &Reader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    document: &mut XmlDocument,
    stack: &mut Vec<(Node, HashMap<String, String>)>,
    string_bytes: &mut usize,
    xml_version: XmlVersion,
    empty: bool,
) -> Result<()> {
    if stack.len() >= MAX_WEB_EXTENSION_XML_DEPTH {
        return invalid(format!(
            "web extension XML depth exceeds {MAX_WEB_EXTENSION_XML_DEPTH}"
        ));
    }
    let mut namespaces = stack
        .last()
        .map(|(_, namespaces)| namespaces.clone())
        .unwrap_or_default();
    namespaces.insert(
        "xml".into(),
        "http://www.w3.org/XML/1998/namespace".into(),
    );
    let mut raw_attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(xml_version, reader.decoder())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        *string_bytes = string_bytes.saturating_add(name.len() + value.len());
        if *string_bytes > MAX_WEB_EXTENSION_STRING_BYTES {
            return invalid("web extension decoded strings exceed allocation limit".into());
        }
        if name == "xmlns" {
            namespaces.insert(String::new(), value);
        } else if let Some(prefix) = name.strip_prefix("xmlns:") {
            namespaces.insert(prefix.to_owned(), value);
        } else {
            raw_attributes.push((name, value));
        }
    }
    let element_name = element.name();
    let raw_name = std::str::from_utf8(element_name.as_ref())
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    let (prefix, local_name) = split_qname(raw_name);
    let namespace = namespaces.get(prefix).cloned().ok_or_else(|| {
        OoxmlError::InvalidFormat(format!("unbound XML namespace prefix '{prefix}'"))
    })?;
    *string_bytes = string_bytes.saturating_add(namespace.len() + local_name.len());
    let mut attributes = Vec::with_capacity(raw_attributes.len());
    let mut seen = HashSet::new();
    for (raw_name, value) in raw_attributes {
        let (prefix, local_name) = split_qname(&raw_name);
        let namespace = if prefix.is_empty() {
            String::new()
        } else {
            namespaces.get(prefix).cloned().ok_or_else(|| {
                OoxmlError::InvalidFormat(format!("unbound attribute prefix '{prefix}'"))
            })?
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
    let node = Node {
        namespace,
        local_name: local_name.to_owned(),
        attributes,
        children: Vec::new(),
    };
    if empty {
        attach_node(document, stack, node)?;
    } else {
        stack.push((node, namespaces));
    }
    Ok(())
}

fn attach_node(
    document: &mut XmlDocument,
    stack: &mut [(Node, HashMap<String, String>)],
    node: Node,
) -> Result<()> {
    if let Some((parent, _)) = stack.last_mut() {
        parent.children.push(node);
    } else if document.root.replace(node).is_some() {
        return invalid("multiple XML root elements".into());
    }
    Ok(())
}

fn split_qname(name: &str) -> (&str, &str) {
    name.split_once(':').unwrap_or(("", name))
}

fn element_children<'a>(node: &'a Node) -> Vec<&'a Node> {
    node.children.iter().collect()
}

fn require_name(node: &Node, namespace: &str, local_name: &str) -> Result<()> {
    if node.namespace == namespace && node.local_name == local_name {
        Ok(())
    } else {
        invalid(format!(
            "expected {{{namespace}}}{local_name}, got {{{}}}{}",
            node.namespace, node.local_name
        ))
    }
}

fn attr<'a>(node: &'a Node, namespace: &str, local_name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| {
            attribute.namespace == namespace && attribute.local_name == local_name
        })
        .map(|attribute| attribute.value.as_str())
}

fn required_attr<'a>(node: &'a Node, namespace: &str, local_name: &str) -> Result<&'a str> {
    attr(node, namespace, local_name).ok_or_else(|| {
        OoxmlError::InvalidFormat(format!(
            "{} requires attribute {{{namespace}}}{local_name}",
            node.local_name
        ))
    })
}

fn relationship_attr<'a>(node: &'a Node, local_name: &str) -> Result<Option<&'a str>> {
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

fn is_relationship_namespace(namespace: &str) -> bool {
    matches!(
        namespace,
        TRANSITIONAL_RELATIONSHIPS_NAMESPACE | STRICT_RELATIONSHIPS_NAMESPACE
    )
}

fn optional_bool_attr(node: &Node, namespace: &str, local_name: &str) -> Result<Option<bool>> {
    attr(node, namespace, local_name).map(parse_bool).transpose()
}

fn reject_unknown_attributes(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
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

fn is_next(children: &[&Node], position: usize, namespace: &str, local_name: &str) -> bool {
    children.get(position).is_some_and(|child| {
        child.namespace == namespace && child.local_name == local_name
    })
}

fn next_required<'a>(
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

fn ensure_consumed(children: &[&Node], position: usize, parent: &str) -> Result<()> {
    if position == children.len() {
        Ok(())
    } else {
        invalid(format!(
            "unexpected child {} in {parent}",
            children[position].local_name
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::PackURI;
    use litchi_opc::part::XmlPart;

    const LO_OMEX_DOCX: &[u8] = include_bytes!(
        "../../../3rdparty/libreoffice-core/sw/qa/extras/ooxmlimport/data/n820504.docx"
    );
    const LO_REGISTRY_DOCX: &[u8] = include_bytes!(
        "../../../3rdparty/libreoffice-core/sw/qa/core/objectpositioning/data/do-not-capture-draw-objs-on-page-draw-wrap-none.docx"
    );

    #[test]
    fn loads_libreoffice_omex_and_registry_fixtures_inertly() {
        let omex = OpcPackage::from_bytes(LO_OMEX_DOCX).unwrap();
        let panes = load_web_extension_task_panes(&omex).unwrap().unwrap();
        assert_eq!(panes.panes.len(), 1);
        assert_eq!(panes.panes[0].web_extension.reference.store_type, WebExtensionStoreType::Omex);
        assert!(panes.panes[0].visible);

        let registry = OpcPackage::from_bytes(LO_REGISTRY_DOCX).unwrap();
        let panes = load_web_extension_task_panes(&registry).unwrap().unwrap();
        assert_eq!(panes.panes[0].web_extension.reference.store_type, WebExtensionStoreType::Registry);
        assert!(!panes.panes[0].visible);
    }

    #[test]
    fn strict_writer_is_deterministic_and_round_trips() {
        let extension = sample_extension();
        let first = write_web_extension(&extension, OoxmlConformance::Strict).unwrap();
        let second = write_web_extension(&extension, OoxmlConformance::Strict).unwrap();
        assert_eq!(first, second);
        assert!(std::str::from_utf8(&first)
            .unwrap()
            .contains(STRICT_RELATIONSHIPS_NAMESPACE));
        assert_eq!(parse_web_extension(&first).unwrap(), extension);
    }

    #[test]
    fn accepts_mce_alternate_content_and_strict_relationship_attributes() {
        let xml = format!(r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" xmlns:r="{STRICT_RELATIONSHIPS_NAMESPACE}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" id="x"><we:reference id="a" version="1"/><mc:AlternateContent><mc:Choice Requires="we"><we:alternateReferences/></mc:Choice><mc:Fallback/></mc:AlternateContent><we:properties/><we:bindings/><we:snapshot r:embed="rId1"/></we:webextension>"#);
        let extension = parse_web_extension(xml.as_bytes()).unwrap();
        assert_eq!(extension.snapshot.unwrap().embedded_relationship_id.as_deref(), Some("rId1"));
    }

    #[test]
    fn rejects_dtd_bad_order_bad_store_and_nonfinite_width() {
        assert!(parse_web_extension(br#"<!DOCTYPE x><x/>"#).is_err());
        let bad_order = format!(r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" id="x"><we:properties/><we:reference id="a" version="1"/><we:bindings/></we:webextension>"#);
        assert!(parse_web_extension(bad_order.as_bytes()).is_err());
        let bad_store = format!(r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" id="x"><we:reference id="a" version="1" storeType="Network"/><we:properties/><we:bindings/></we:webextension>"#);
        assert!(parse_web_extension(bad_store.as_bytes()).is_err());
        let bad_width = format!(r#"<wetp:taskpanes xmlns:wetp="{TASK_PANES_NAMESPACE}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS_NAMESPACE}"><wetp:taskpane dockstate="right" visibility="1" width="NaN" row="0"><wetp:webextensionref r:id="rId1"/></wetp:taskpane></wetp:taskpanes>"#);
        assert!(parse_task_panes(bad_width.as_bytes()).is_err());
    }

    #[test]
    fn enforces_input_and_list_caps() {
        assert!(parse_web_extension(&vec![b' '; MAX_WEB_EXTENSION_XML_BYTES + 1]).is_err());
        let mut model = WebExtensionTaskPanes::default();
        model.panes.resize_with(MAX_WEB_EXTENSION_ITEMS + 1, || WebExtensionTaskPane {
            dock_state: "right".into(),
            visible: false,
            width: 320.0,
            row: 0,
            locked: false,
            relationship_id: "rId1".into(),
            web_extension: sample_extension(),
        });
        assert!(write_task_panes(&model, OoxmlConformance::Transitional).is_err());
    }

    #[test]
    fn rejects_external_wrong_content_type_and_dangling_package_graphs() {
        let external = synthetic_package(true, WEB_EXTENSION_CONTENT_TYPE, "rId1");
        assert!(load_web_extension_task_panes(&external).is_err());

        let wrong_type = synthetic_package(false, "application/xml", "rId1");
        assert!(matches!(
            load_web_extension_task_panes(&wrong_type),
            Err(OoxmlError::InvalidContentType { .. })
        ));

        let dangling = synthetic_package(false, WEB_EXTENSION_CONTENT_TYPE, "missing");
        assert!(load_web_extension_task_panes(&dangling).is_err());
    }

    fn synthetic_package(
        external_extension: bool,
        extension_content_type: &str,
        pane_relationship_id: &str,
    ) -> OpcPackage {
        let task_panes_xml = format!(
            r#"<wetp:taskpanes xmlns:wetp="{TASK_PANES_NAMESPACE}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS_NAMESPACE}"><wetp:taskpane dockstate="right" visibility="0" width="320" row="0"><wetp:webextensionref r:id="{pane_relationship_id}"/></wetp:taskpane></wetp:taskpanes>"#
        );
        let extension_xml = write_web_extension(
            &sample_extension(),
            OoxmlConformance::Transitional,
        )
        .unwrap();
        let mut package = OpcPackage::new();
        package.rels_mut().add_relationship(
            TASK_PANES_RELATIONSHIP_TYPE.into(),
            "word/webextensions/taskpanes.xml".into(),
            "rIdTaskPanes".into(),
            false,
        );
        let mut task_panes_part = XmlPart::new(
            PackURI::new("/word/webextensions/taskpanes.xml").unwrap(),
            TASK_PANES_CONTENT_TYPE.into(),
            task_panes_xml.into_bytes(),
        );
        task_panes_part.rels_mut().add_relationship(
            WEB_EXTENSION_RELATIONSHIP_TYPE.into(),
            if external_extension {
                "https://example.invalid/add-in".into()
            } else {
                "webextension1.xml".into()
            },
            "rId1".into(),
            external_extension,
        );
        package.add_part(Box::new(task_panes_part));
        if !external_extension {
            package.add_part(Box::new(XmlPart::new(
                PackURI::new("/word/webextensions/webextension1.xml").unwrap(),
                extension_content_type.into(),
                extension_xml,
            )));
        }
        package
    }

    fn sample_extension() -> WebExtension {
        WebExtension {
            id: "{00000000-0000-0000-0000-000000000001}".into(),
            frozen: true,
            reference: WebExtensionStoreReference {
                id: "wa1".into(),
                version: "1.0.0.0".into(),
                store: Some("en-us".into()),
                store_type: WebExtensionStoreType::Omex,
            },
            alternate_references: vec![],
            properties: vec![WebExtensionProperty {
                name: "Office.AutoShowTaskpaneWithDocument".into(),
                value: "false".into(),
            }],
            bindings: vec![WebExtensionBinding {
                id: "binding-1".into(),
                binding_type: "matrix".into(),
                application_reference: "app-ref".into(),
            }],
            snapshot: Some(WebExtensionSnapshot::default()),
        }
    }
}
