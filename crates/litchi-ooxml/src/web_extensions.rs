//! Inert Office Web Extension and persisted task-pane metadata.
//!
//! This module implements the package structures defined by MS-OWEXML. It
//! intentionally does not locate add-ins, contact catalog providers, load
//! manifests, resolve linked content, or execute scripts/custom functions.

use crate::error::{OoxmlError, Result};
use litchi_ooxml_common::{MceCapabilities, MceLimits, process_markup_compatibility};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
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
pub const DRAWINGML_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub const STRICT_DRAWINGML_NAMESPACE: &str = "http://purl.oclc.org/ooxml/drawingml/main";

pub const TASK_PANES_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes";
pub const WEB_EXTENSION_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/webextension";
pub const TASK_PANES_CONTENT_TYPE: &str = "application/vnd.ms-office.webextensiontaskpanes+xml";
pub const WEB_EXTENSION_CONTENT_TYPE: &str = "application/vnd.ms-office.webextension+xml";

const IMAGE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const STRICT_IMAGE_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image";

/// Maximum bytes accepted for either XML part before MCE processing.
pub const MAX_WEB_EXTENSION_XML_BYTES: usize = 4 * 1024 * 1024;
/// Maximum XML nesting depth.
pub const MAX_WEB_EXTENSION_XML_DEPTH: usize = 128;
/// Maximum element count in one web-extension XML part or authored fragment.
pub const MAX_WEB_EXTENSION_XML_NODES: usize = 65_536;
/// Maximum task panes, alternate references, properties, or bindings per list.
pub const MAX_WEB_EXTENSION_ITEMS: usize = 4096;
/// Maximum aggregate decoded string bytes in one XML part.
pub const MAX_WEB_EXTENSION_STRING_BYTES: usize = 8 * 1024 * 1024;
/// Maximum bytes accepted for one inert snapshot image.
pub const MAX_WEB_EXTENSION_SNAPSHOT_BYTES: usize = 64 * 1024 * 1024;
/// Maximum aggregate snapshot bytes retained from one package.
pub const MAX_WEB_EXTENSION_TOTAL_SNAPSHOT_BYTES: usize = 256 * 1024 * 1024;

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

    fn image_relationship_type(self) -> &'static str {
        match self {
            Self::Transitional => IMAGE_RELATIONSHIP_TYPE,
            Self::Strict => STRICT_IMAGE_RELATIONSHIP_TYPE,
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
    pub extension_list: Option<WebExtensionExtensionList>,
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
    pub extension_list: Option<WebExtensionExtensionList>,
}

/// Namespace dialect of an MS-OWEXML extension-list element.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WebExtensionExtensionListKind {
    WebExtension,
    TaskPane,
    DrawingMl,
    StrictDrawingMl,
}

impl WebExtensionExtensionListKind {
    pub fn namespace(self) -> &'static str {
        match self {
            Self::WebExtension => WEB_EXTENSION_NAMESPACE,
            Self::TaskPane => TASK_PANES_NAMESPACE,
            Self::DrawingMl => DRAWINGML_NAMESPACE,
            Self::StrictDrawingMl => STRICT_DRAWINGML_NAMESPACE,
        }
    }

    fn from_namespace(namespace: &str) -> Result<Self> {
        match namespace {
            WEB_EXTENSION_NAMESPACE => Ok(Self::WebExtension),
            TASK_PANES_NAMESPACE => Ok(Self::TaskPane),
            DRAWINGML_NAMESPACE => Ok(Self::DrawingMl),
            STRICT_DRAWINGML_NAMESPACE => Ok(Self::StrictDrawingMl),
            _ => invalid(format!(
                "invalid web extension extLst namespace '{namespace}'"
            )),
        }
    }
}

/// A bounded, self-contained, inert `extLst` fragment.
///
/// Unknown extension payloads are retained without interpretation or resource
/// resolution. Namespace declarations inherited by the source fragment are
/// materialized on its root so it remains valid when authored elsewhere.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WebExtensionExtensionList {
    kind: WebExtensionExtensionListKind,
    xml: String,
}

impl WebExtensionExtensionList {
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_WEB_EXTENSION_XML_BYTES {
            return invalid(format!(
                "web extension extLst XML exceeds {MAX_WEB_EXTENSION_XML_BYTES} bytes"
            ));
        }
        let document = parse_xml(xml)?;
        Self::from_node(document.root()?, &document)
    }

    pub fn kind(&self) -> WebExtensionExtensionListKind {
        self.kind
    }

    pub fn as_xml(&self) -> &[u8] {
        self.xml.as_bytes()
    }

    pub fn xml(&self) -> &str {
        &self.xml
    }

    fn from_node(node: &Node, document: &XmlDocument) -> Result<Self> {
        if node.local_name != "extLst" {
            return invalid(format!(
                "web extension extension fragment root must be extLst, got {}",
                node.local_name
            ));
        }
        reject_unknown_attributes(node, &[])?;
        let kind = WebExtensionExtensionListKind::from_namespace(&node.namespace)?;
        Ok(Self {
            kind,
            xml: document.self_contained_fragment(node)?,
        })
    }
}

/// Compression state of a DrawingML `CT_Blip`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WebExtensionBlipCompression {
    Email,
    Screen,
    Print,
    HighQualityPrint,
    None,
}

impl WebExtensionBlipCompression {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Email => "email",
            Self::Screen => "screen",
            Self::Print => "print",
            Self::HighQualityPrint => "hqprint",
            Self::None => "none",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "email" => Ok(Self::Email),
            "screen" => Ok(Self::Screen),
            "print" => Ok(Self::Print),
            "hqprint" => Ok(Self::HighQualityPrint),
            "none" => Ok(Self::None),
            _ => invalid(format!("invalid snapshot compression state '{value}'")),
        }
    }
}

/// Closed effect-element choice allowed by DrawingML `CT_Blip`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WebExtensionSnapshotEffectKind {
    AlphaBiLevel,
    AlphaCeiling,
    AlphaFloor,
    AlphaInverse,
    AlphaModulate,
    AlphaModulateFixed,
    AlphaReplace,
    BiLevel,
    Blur,
    ColorChange,
    ColorReplace,
    Duotone,
    FillOverlay,
    Grayscale,
    HueSaturationLuminance,
    Luminance,
    Tint,
}

impl WebExtensionSnapshotEffectKind {
    pub fn local_name(self) -> &'static str {
        match self {
            Self::AlphaBiLevel => "alphaBiLevel",
            Self::AlphaCeiling => "alphaCeiling",
            Self::AlphaFloor => "alphaFloor",
            Self::AlphaInverse => "alphaInv",
            Self::AlphaModulate => "alphaMod",
            Self::AlphaModulateFixed => "alphaModFix",
            Self::AlphaReplace => "alphaRepl",
            Self::BiLevel => "biLevel",
            Self::Blur => "blur",
            Self::ColorChange => "clrChange",
            Self::ColorReplace => "clrRepl",
            Self::Duotone => "duotone",
            Self::FillOverlay => "fillOverlay",
            Self::Grayscale => "grayscl",
            Self::HueSaturationLuminance => "hsl",
            Self::Luminance => "lum",
            Self::Tint => "tint",
        }
    }

    fn parse(local_name: &str) -> Result<Self> {
        match local_name {
            "alphaBiLevel" => Ok(Self::AlphaBiLevel),
            "alphaCeiling" => Ok(Self::AlphaCeiling),
            "alphaFloor" => Ok(Self::AlphaFloor),
            "alphaInv" => Ok(Self::AlphaInverse),
            "alphaMod" => Ok(Self::AlphaModulate),
            "alphaModFix" => Ok(Self::AlphaModulateFixed),
            "alphaRepl" => Ok(Self::AlphaReplace),
            "biLevel" => Ok(Self::BiLevel),
            "blur" => Ok(Self::Blur),
            "clrChange" => Ok(Self::ColorChange),
            "clrRepl" => Ok(Self::ColorReplace),
            "duotone" => Ok(Self::Duotone),
            "fillOverlay" => Ok(Self::FillOverlay),
            "grayscl" => Ok(Self::Grayscale),
            "hsl" => Ok(Self::HueSaturationLuminance),
            "lum" => Ok(Self::Luminance),
            "tint" => Ok(Self::Tint),
            _ => invalid(format!("invalid snapshot effect '{local_name}'")),
        }
    }
}

/// A validated, inert DrawingML effect subtree.
///
/// The subtree is retained as canonical XML. It is never interpreted as
/// executable content, and construction rejects text, CDATA, DTDs, excessive
/// depth, and roots outside the closed `CT_Blip` effect choice.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WebExtensionSnapshotEffect {
    kind: WebExtensionSnapshotEffectKind,
    xml: String,
}

impl WebExtensionSnapshotEffect {
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_WEB_EXTENSION_XML_BYTES {
            return invalid(format!(
                "snapshot effect XML exceeds {MAX_WEB_EXTENSION_XML_BYTES} bytes"
            ));
        }
        let document = parse_xml(xml)?;
        Self::from_node(document.root()?)
    }

    pub fn kind(&self) -> WebExtensionSnapshotEffectKind {
        self.kind
    }

    pub fn xml(&self) -> &str {
        &self.xml
    }

    fn from_node(node: &Node) -> Result<Self> {
        if !is_drawingml_namespace(&node.namespace) {
            return invalid(format!(
                "snapshot effect {} has invalid namespace '{}'",
                node.local_name, node.namespace
            ));
        }
        let kind = WebExtensionSnapshotEffectKind::parse(&node.local_name)?;
        Ok(Self {
            kind,
            xml: canonical_node_xml(node),
        })
    }
}

/// DrawingML `CT_Blip` metadata used by a web-extension snapshot.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct WebExtensionSnapshot {
    pub embedded_relationship_id: Option<String>,
    pub linked_relationship_id: Option<String>,
    pub compression_state: Option<WebExtensionBlipCompression>,
    pub effects: Vec<WebExtensionSnapshotEffect>,
    pub extension_list: Option<WebExtensionExtensionList>,
}

/// One inert image relationship owned by a web-extension snapshot.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WebExtensionSnapshotResource {
    pub relationship_id: String,
    pub target: WebExtensionSnapshotTarget,
}

/// Internal image bytes or an external linked image target.
///
/// External targets are retained as strings and are never fetched.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum WebExtensionSnapshotTarget {
    Internal {
        part_name: PackURI,
        content_type: String,
        data: Vec<u8>,
    },
    External {
        target: String,
    },
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
    pub extension_list: Option<WebExtensionExtensionList>,
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
    pub snapshot_resources: Vec<WebExtensionSnapshotResource>,
    pub extension_list: Option<WebExtensionExtensionList>,
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
        enforce_count("alternate reference", refs.len())?;
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
            .map(WebExtensionBlipCompression::parse)
            .transpose()?;
        let snapshot_children = element_children(node);
        enforce_count("snapshot effect", snapshot_children.len())?;
        let mut effects = Vec::with_capacity(snapshot_children.len());
        let mut extension_list = None;
        for (index, child) in snapshot_children.iter().enumerate() {
            if is_drawingml_namespace(&child.namespace) && child.local_name == "extLst" {
                if index + 1 != snapshot_children.len() {
                    return invalid("snapshot extLst must be the final child".into());
                }
                extension_list = Some(WebExtensionExtensionList::from_node(child, &document)?);
                continue;
            }
            effects.push(WebExtensionSnapshotEffect::from_node(child)?);
        }
        Some(WebExtensionSnapshot {
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
        let value = WebExtensionExtensionList::from_node(children[position], &document)?;
        position += 1;
        Some(value)
    } else {
        None
    };
    ensure_consumed(&children, position, "webextension")?;

    Ok(WebExtension {
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
pub fn parse_task_panes(xml: &[u8]) -> Result<Vec<ParsedTaskPane>> {
    let document = parse_mce_xml(xml, &[TASK_PANES_NAMESPACE, WEB_EXTENSION_NAMESPACE])?;
    let root = document.root()?;
    require_name(root, TASK_PANES_NAMESPACE, "taskpanes")?;
    reject_unknown_attributes(root, &[])?;
    let children = element_children(root);
    enforce_count("task pane", children.len())?;
    children
        .into_iter()
        .map(|node| parse_task_pane(node, &document))
        .collect()
}

#[derive(Debug, Clone, PartialEq)]
pub struct ParsedTaskPane {
    pub dock_state: String,
    pub visible: bool,
    pub width: f64,
    pub row: u32,
    pub locked: bool,
    pub relationship_id: String,
    pub extension_list: Option<WebExtensionExtensionList>,
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
    if referenced_ids.len() != parsed_panes.len() {
        return invalid("task panes contain duplicate relationship IDs".into());
    }
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
    let mut total_snapshot_bytes = 0usize;
    let mut extension_names = HashSet::new();
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
        if !extension_names.insert(extension_name.clone()) {
            return invalid(format!(
                "multiple task panes target web extension part '{}'",
                extension_name.as_str()
            ));
        }
        let extension_part = package.get_part(&extension_name).map_err(|error| {
            OoxmlError::PartNotFound(format!(
                "web extension part '{}': {error}",
                extension_name.as_str()
            ))
        })?;
        require_content_type(extension_part, WEB_EXTENSION_CONTENT_TYPE)?;
        let web_extension = parse_web_extension(extension_part.blob())?;
        let snapshot_resources = load_snapshot_resources(
            package,
            extension_part,
            &web_extension,
            &mut total_snapshot_bytes,
        )?;
        panes.push(WebExtensionTaskPane {
            dock_state: pane.dock_state,
            visible: pane.visible,
            width: pane.width,
            row: pane.row,
            locked: pane.locked,
            relationship_id: pane.relationship_id,
            web_extension,
            snapshot_resources,
            extension_list: pane.extension_list,
        });
    }
    Ok(Some(WebExtensionTaskPanes { panes }))
}

/// Create or replace the package-level persisted task-pane graph.
///
/// Add-in references, bindings, properties, and snapshot resources are stored
/// as inert data. External snapshot links are never contacted.
pub fn store_web_extension_task_panes(
    package: &mut OpcPackage,
    task_panes: &WebExtensionTaskPanes,
    conformance: OoxmlConformance,
) -> Result<()> {
    let task_panes_xml = write_task_panes(task_panes, conformance)?;
    let existing = existing_web_extension_graph(package)?;
    let task_panes_name = existing
        .as_ref()
        .map(|graph| graph.task_panes_name.clone())
        .map_or_else(|| next_task_panes_part_name(package), Ok)?;
    let mut reserved = HashSet::new();
    reserved.insert(task_panes_name.clone());
    let mut planned = Vec::with_capacity(task_panes.panes.len() + 1);
    let mut task_relationships = Vec::with_capacity(task_panes.panes.len());
    let mut total_snapshot_bytes = 0usize;
    let mut counted_snapshot_parts = HashSet::new();
    let existing_extensions = existing
        .as_ref()
        .map(|graph| &graph.extensions_by_relationship);

    for (index, pane) in task_panes.panes.iter().enumerate() {
        let extension_name = match existing_extensions
            .and_then(|extensions| extensions.get(&pane.relationship_id))
        {
            Some(name) => name.clone(),
            None => next_web_extension_part_name(package, &reserved, index + 1)?,
        };
        if !reserved.insert(extension_name.clone()) {
            return invalid(format!(
                "multiple task panes target web extension part '{}'",
                extension_name.as_str()
            ));
        }
        let extension_xml = write_web_extension(&pane.web_extension, conformance)?;
        let mut relationships = Vec::with_capacity(pane.snapshot_resources.len());
        for resource in &pane.snapshot_resources {
            let (target, external) = match &resource.target {
                WebExtensionSnapshotTarget::Internal {
                    part_name,
                    content_type,
                    data,
                } => {
                    reserved.insert(part_name.clone());
                    if counted_snapshot_parts.insert(part_name.clone()) {
                        total_snapshot_bytes = total_snapshot_bytes
                            .checked_add(data.len())
                            .ok_or_else(|| {
                                OoxmlError::InvalidFormat(
                                    "aggregate snapshot byte count overflow".into(),
                                )
                            })?;
                        if total_snapshot_bytes > MAX_WEB_EXTENSION_TOTAL_SNAPSHOT_BYTES {
                            return invalid(format!(
                                "aggregate snapshot images exceed {MAX_WEB_EXTENSION_TOTAL_SNAPSHOT_BYTES} bytes"
                            ));
                        }
                    }
                    add_or_match_planned_part(
                        &mut planned,
                        PlannedPart {
                            name: part_name.clone(),
                            content_type: content_type.clone(),
                            data: data.clone(),
                            relationships: Vec::new(),
                        },
                    )?;
                    (part_name.relative_ref(extension_name.base_uri()), false)
                },
                WebExtensionSnapshotTarget::External { target } => (target.clone(), true),
            };
            relationships.push(PlannedRelationship {
                id: resource.relationship_id.clone(),
                relationship_type: conformance.image_relationship_type().into(),
                target,
                external,
            });
        }
        add_or_match_planned_part(
            &mut planned,
            PlannedPart {
                name: extension_name.clone(),
                content_type: WEB_EXTENSION_CONTENT_TYPE.into(),
                data: extension_xml,
                relationships,
            },
        )?;
        task_relationships.push(PlannedRelationship {
            id: pane.relationship_id.clone(),
            relationship_type: WEB_EXTENSION_RELATIONSHIP_TYPE.into(),
            target: extension_name.relative_ref(task_panes_name.base_uri()),
            external: false,
        });
    }
    add_or_match_planned_part(
        &mut planned,
        PlannedPart {
            name: task_panes_name.clone(),
            content_type: TASK_PANES_CONTENT_TYPE.into(),
            data: task_panes_xml,
            relationships: task_relationships,
        },
    )?;

    let old_parts = existing
        .as_ref()
        .map_or(&[][..], |graph| graph.owned_parts.as_slice());
    preflight_planned_parts(package, &planned, old_parts)?;
    install_planned_parts(package, planned, old_parts)?;

    let root_relationship_id = existing
        .as_ref()
        .map(|graph| graph.root_relationship_id.clone())
        .unwrap_or_else(|| next_package_relationship_id(package));
    if existing.is_some() {
        package.rels_mut().remove(&root_relationship_id);
    }
    package.rels_mut().add_relationship(
        TASK_PANES_RELATIONSHIP_TYPE.into(),
        task_panes_name.as_str().trim_start_matches('/').into(),
        root_relationship_id,
        false,
    );

    if let Some(existing) = existing {
        remove_unreferenced_parts(package, &existing.owned_parts, &reserved);
    }
    let _ = package.clear_digital_signatures();
    Ok(())
}

/// Remove the package-level task-pane relationship and graph.
///
/// Parts still referenced elsewhere remain in the package.
pub fn remove_web_extension_task_panes(package: &mut OpcPackage) -> Result<bool> {
    let Some(existing) = existing_web_extension_graph(package)? else {
        return Ok(false);
    };
    package.rels_mut().remove(&existing.root_relationship_id);
    remove_unreferenced_parts(package, &existing.owned_parts, &HashSet::new());
    let _ = package.clear_digital_signatures();
    Ok(true)
}

#[derive(Debug)]
struct ExistingWebExtensionGraph {
    root_relationship_id: String,
    task_panes_name: PackURI,
    extensions_by_relationship: HashMap<String, PackURI>,
    owned_parts: Vec<PackURI>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct PlannedRelationship {
    id: String,
    relationship_type: String,
    target: String,
    external: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct PlannedPart {
    name: PackURI,
    content_type: String,
    data: Vec<u8>,
    relationships: Vec<PlannedRelationship>,
}

fn existing_web_extension_graph(package: &OpcPackage) -> Result<Option<ExistingWebExtensionGraph>> {
    let Some(loaded) = load_web_extension_task_panes(package)? else {
        return Ok(None);
    };
    let relationship = package
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP_TYPE)
        .ok_or_else(|| {
            OoxmlError::InvalidRelationship("loaded task panes have no package relationship".into())
        })?;
    let task_panes_name = relationship.target_partname().map_err(|error| {
        OoxmlError::InvalidRelationship(format!("invalid task-pane target: {error}"))
    })?;
    let task_panes_part = package.get_part(&task_panes_name)?;
    let mut extensions_by_relationship = HashMap::with_capacity(loaded.panes.len());
    let mut owned = HashSet::new();
    owned.insert(task_panes_name.clone());
    for pane in loaded.panes {
        let child_relationship = task_panes_part
            .rels()
            .get(&pane.relationship_id)
            .ok_or_else(|| {
                OoxmlError::InvalidRelationship(format!(
                    "task pane references missing relationship '{}'",
                    pane.relationship_id
                ))
            })?;
        let extension_name = child_relationship.target_partname().map_err(|error| {
            OoxmlError::InvalidRelationship(format!(
                "invalid web extension target '{}': {error}",
                pane.relationship_id
            ))
        })?;
        if extensions_by_relationship
            .insert(pane.relationship_id.clone(), extension_name.clone())
            .is_some()
        {
            return invalid(format!(
                "duplicate task-pane relationship ID '{}'",
                pane.relationship_id
            ));
        }
        owned.insert(extension_name);
        for resource in pane.snapshot_resources {
            if let WebExtensionSnapshotTarget::Internal { part_name, .. } = resource.target {
                owned.insert(part_name);
            }
        }
    }
    let mut owned_parts: Vec<_> = owned.into_iter().collect();
    owned_parts.sort_by(|left, right| left.as_str().cmp(right.as_str()));
    Ok(Some(ExistingWebExtensionGraph {
        root_relationship_id: relationship.r_id().to_owned(),
        task_panes_name,
        extensions_by_relationship,
        owned_parts,
    }))
}

fn next_web_extension_part_name(
    package: &OpcPackage,
    reserved: &HashSet<PackURI>,
    preferred_index: usize,
) -> Result<PackURI> {
    for offset in 0..=MAX_WEB_EXTENSION_ITEMS {
        let index = preferred_index
            .checked_add(offset)
            .ok_or_else(|| OoxmlError::InvalidFormat("web extension part index overflow".into()))?;
        let candidate = PackURI::new(format!("/webextensions/webextension{index}.xml"))
            .map_err(OoxmlError::InvalidUri)?;
        if !reserved.contains(&candidate)
            && package
                .iter_parts()
                .all(|part| part.partname() != &candidate)
        {
            return Ok(candidate);
        }
    }
    invalid("no free web extension part name".into())
}

fn next_task_panes_part_name(package: &OpcPackage) -> Result<PackURI> {
    for index in 1..=MAX_WEB_EXTENSION_ITEMS + 1 {
        let suffix = if index == 1 {
            String::new()
        } else {
            index.to_string()
        };
        let candidate = PackURI::new(format!("/webextensions/taskpanes{suffix}.xml"))
            .map_err(OoxmlError::InvalidUri)?;
        if package
            .iter_parts()
            .all(|part| part.partname() != &candidate)
        {
            return Ok(candidate);
        }
    }
    invalid("no free task-pane part name".into())
}

fn next_package_relationship_id(package: &OpcPackage) -> String {
    for index in 1..=u32::MAX {
        let candidate = format!("rIdWebExtensionTaskPanes{index}");
        if package.rels().get(&candidate).is_none() {
            return candidate;
        }
    }
    unreachable!("finite relationship collection cannot consume every u32 ID")
}

fn add_or_match_planned_part(planned: &mut Vec<PlannedPart>, part: PlannedPart) -> Result<()> {
    if let Some(existing) = planned.iter().find(|existing| existing.name == part.name) {
        if existing == &part {
            return Ok(());
        }
        return invalid(format!(
            "conflicting authored resources target '{}'",
            part.name.as_str()
        ));
    }
    planned.push(part);
    Ok(())
}

fn preflight_planned_parts(
    package: &OpcPackage,
    planned: &[PlannedPart],
    old_parts: &[PackURI],
) -> Result<()> {
    let old_parts: HashSet<_> = old_parts.iter().collect();
    for part in planned {
        if let Ok(existing) = package.get_part(&part.name) {
            if !old_parts.contains(&part.name) {
                return invalid(format!(
                    "authored web extension part '{}' already exists outside the replaced graph",
                    part.name.as_str()
                ));
            }
            if existing.content_type() != part.content_type {
                return invalid(format!(
                    "cannot change content type of existing part '{}'",
                    part.name.as_str()
                ));
            }
            if part.content_type.starts_with("image/")
                && existing.blob() != part.data
                && package_part_is_referenced_outside(package, &part.name, &old_parts)
            {
                return invalid(format!(
                    "cannot replace shared snapshot resource '{}'",
                    part.name.as_str()
                ));
            }
        }
    }
    Ok(())
}

fn install_planned_parts(
    package: &mut OpcPackage,
    planned: Vec<PlannedPart>,
    old_parts: &[PackURI],
) -> Result<()> {
    let old_parts: HashSet<_> = old_parts.iter().collect();
    let mut added = Vec::new();
    let mut replacements = Vec::new();
    for part in planned {
        if old_parts.contains(&part.name) {
            replacements.push(part);
            continue;
        }
        let name = part.name.clone();
        let value = planned_blob_part(part);
        if let Err(error) = package.try_add_part(Box::new(value)) {
            for name in added {
                package.remove_part(&name);
            }
            return Err(error.into());
        }
        added.push(name);
    }
    for part in replacements {
        let existing = package.get_part_mut(&part.name)?;
        existing.set_blob(part.data);
        let relationship_ids: Vec<_> = existing
            .rels()
            .iter()
            .map(|relationship| relationship.r_id().to_owned())
            .collect();
        for id in relationship_ids {
            existing.rels_mut().remove(&id);
        }
        for relationship in part.relationships {
            existing.rels_mut().add_relationship(
                relationship.relationship_type,
                relationship.target,
                relationship.id,
                relationship.external,
            );
        }
    }
    Ok(())
}

fn planned_blob_part(part: PlannedPart) -> BlobPart {
    let mut value = BlobPart::new(part.name, part.content_type, part.data);
    for relationship in part.relationships {
        value.rels_mut().add_relationship(
            relationship.relationship_type,
            relationship.target,
            relationship.id,
            relationship.external,
        );
    }
    value
}

fn remove_unreferenced_parts(
    package: &mut OpcPackage,
    candidates: &[PackURI],
    retained: &HashSet<PackURI>,
) {
    loop {
        let removable: Vec<_> = candidates
            .iter()
            .filter(|name| {
                !retained.contains(*name)
                    && package.get_part(name).is_ok()
                    && !package_part_is_referenced(package, name)
            })
            .cloned()
            .collect();
        if removable.is_empty() {
            break;
        }
        for name in removable {
            package.remove_part(&name);
        }
    }
}

fn package_part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|name| name == *target)
    }) || package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|name| name == *target)
        })
    })
}

fn package_part_is_referenced_outside(
    package: &OpcPackage,
    target: &PackURI,
    allowed_sources: &HashSet<&PackURI>,
) -> bool {
    package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|name| name == *target)
    }) || package.iter_parts().any(|part| {
        !allowed_sources.contains(part.partname())
            && part.rels().iter().any(|relationship| {
                !relationship.is_external()
                    && relationship
                        .target_partname()
                        .is_ok_and(|name| name == *target)
            })
    })
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
    parse_web_extension(&output)?;
    Ok(output)
}

/// Deterministically serialize task-pane metadata and relationship IDs.
pub fn write_task_panes(
    task_panes: &WebExtensionTaskPanes,
    conformance: OoxmlConformance,
) -> Result<Vec<u8>> {
    enforce_count("task pane", task_panes.panes.len())?;
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
        validate_task_pane(pane)?;
        if !relationship_ids.insert(pane.relationship_id.as_str()) {
            return invalid(format!(
                "duplicate task-pane relationship ID '{}'",
                pane.relationship_id
            ));
        }
        if !extension_ids.insert(pane.web_extension.id.as_str()) {
            return invalid(format!(
                "duplicate web extension instance ID '{}'",
                pane.web_extension.id
            ));
        }
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
        out.push_str("\"/>");
        if let Some(extension_list) = &pane.extension_list {
            out.push_str(extension_list.xml());
        }
        out.push_str("</wetp:taskpane>");
    }
    out.push_str("</wetp:taskpanes>");
    let output = out.into_bytes();
    parse_task_panes(&output)?;
    Ok(output)
}

fn parse_task_pane(node: &Node, document: &XmlDocument) -> Result<ParsedTaskPane> {
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
            && (children[1].namespace != TASK_PANES_NAMESPACE
                || children[1].local_name != "extLst"))
    {
        return invalid("unexpected taskpane child or child order".into());
    }
    let extension_list = children
        .get(1)
        .map(|node| WebExtensionExtensionList::from_node(node, document))
        .transpose()?;
    Ok(ParsedTaskPane {
        dock_state,
        visible,
        width,
        row,
        locked,
        relationship_id,
        extension_list,
    })
}

fn parse_store_reference(
    node: &Node,
    document: &XmlDocument,
) -> Result<WebExtensionStoreReference> {
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
    Ok(WebExtensionStoreReference {
        id: required_attr(node, "", "id")?.to_owned(),
        version: required_attr(node, "", "version")?.to_owned(),
        store: attr(node, "", "store").map(str::to_owned),
        store_type: attr(node, "", "storeType")
            .map(WebExtensionStoreType::parse)
            .transpose()?
            .unwrap_or_default(),
        extension_list: children
            .first()
            .map(|node| WebExtensionExtensionList::from_node(node, document))
            .transpose()?,
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

fn parse_binding(node: &Node, document: &XmlDocument) -> Result<WebExtensionBinding> {
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
    Ok(WebExtensionBinding {
        id: required_attr(node, "", "id")?.to_owned(),
        binding_type: required_attr(node, "", "type")?.to_owned(),
        application_reference: required_attr(node, "", "appref")?.to_owned(),
        extension_list: children
            .first()
            .map(|node| WebExtensionExtensionList::from_node(node, document))
            .transpose()?,
    })
}

fn load_snapshot_resources(
    package: &OpcPackage,
    part: &dyn Part,
    extension: &WebExtension,
    total_snapshot_bytes: &mut usize,
) -> Result<Vec<WebExtensionSnapshotResource>> {
    let mut referenced = HashMap::new();
    if let Some(snapshot) = &extension.snapshot {
        if let Some(id) = &snapshot.embedded_relationship_id {
            if referenced.insert(id.as_str(), false).is_some() {
                return invalid("snapshot embed and link IDs must differ".into());
            }
        }
        if let Some(id) = &snapshot.linked_relationship_id {
            if referenced.insert(id.as_str(), true).is_some() {
                return invalid("snapshot embed and link IDs must differ".into());
            }
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
            resources.push(WebExtensionSnapshotResource {
                relationship_id: relationship.r_id().to_owned(),
                target: WebExtensionSnapshotTarget::External {
                    target: relationship.target_ref().to_owned(),
                },
            });
            continue;
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
        if image.rels().iter().next().is_some() {
            return invalid(format!(
                "snapshot image '{}' must not have relationships",
                image_name.as_str()
            ));
        }
        if image.blob().len() > MAX_WEB_EXTENSION_SNAPSHOT_BYTES {
            return invalid(format!(
                "snapshot image '{}' exceeds {MAX_WEB_EXTENSION_SNAPSHOT_BYTES} bytes",
                image_name.as_str()
            ));
        }
        *total_snapshot_bytes = total_snapshot_bytes
            .checked_add(image.blob().len())
            .ok_or_else(|| {
                OoxmlError::InvalidFormat("aggregate snapshot byte count overflow".into())
            })?;
        if *total_snapshot_bytes > MAX_WEB_EXTENSION_TOTAL_SNAPSHOT_BYTES {
            return invalid(format!(
                "aggregate snapshot images exceed {MAX_WEB_EXTENSION_TOTAL_SNAPSHOT_BYTES} bytes"
            ));
        }
        resources.push(WebExtensionSnapshotResource {
            relationship_id: relationship.r_id().to_owned(),
            target: WebExtensionSnapshotTarget::Internal {
                part_name: image_name,
                content_type: image.content_type().to_owned(),
                data: image.blob().to_vec(),
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
        validate_extension_list(
            binding.extension_list.as_ref(),
            &[WebExtensionExtensionListKind::WebExtension],
        )?;
    }
    if let Some(snapshot) = &extension.snapshot {
        enforce_count("snapshot effect", snapshot.effects.len())?;
        for effect in &snapshot.effects {
            let reparsed = WebExtensionSnapshotEffect::from_xml(effect.xml.as_bytes())?;
            if reparsed.kind != effect.kind {
                return invalid("snapshot effect kind does not match its XML root".into());
            }
        }
        validate_extension_list(
            snapshot.extension_list.as_ref(),
            &[
                WebExtensionExtensionListKind::DrawingMl,
                WebExtensionExtensionListKind::StrictDrawingMl,
            ],
        )?;
    }
    validate_extension_list(
        extension.extension_list.as_ref(),
        &[WebExtensionExtensionListKind::WebExtension],
    )?;
    Ok(())
}

fn validate_store_reference(reference: &WebExtensionStoreReference) -> Result<()> {
    require_nonempty("reference id", &reference.id)?;
    require_nonempty("reference version", &reference.version)?;
    validate_extension_list(
        reference.extension_list.as_ref(),
        &[WebExtensionExtensionListKind::WebExtension],
    )
}

fn validate_task_pane(pane: &WebExtensionTaskPane) -> Result<()> {
    require_nonempty("dock state", &pane.dock_state)?;
    require_nonempty("task-pane relationship id", &pane.relationship_id)?;
    if !pane.width.is_finite() {
        return invalid("task-pane width must be finite".into());
    }
    validate_extension_list(
        pane.extension_list.as_ref(),
        &[WebExtensionExtensionListKind::TaskPane],
    )?;
    validate_model(&pane.web_extension)?;
    validate_snapshot_resources(pane)
}

fn validate_extension_list(
    extension_list: Option<&WebExtensionExtensionList>,
    allowed: &[WebExtensionExtensionListKind],
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
    let reparsed = WebExtensionExtensionList::from_xml(extension_list.as_xml())?;
    if reparsed != *extension_list {
        return invalid("extLst fragment is not a stable self-contained XML tree".into());
    }
    Ok(())
}

fn validate_snapshot_resources(pane: &WebExtensionTaskPane) -> Result<()> {
    let mut expected = HashMap::new();
    if let Some(snapshot) = &pane.web_extension.snapshot {
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
            WebExtensionSnapshotTarget::Internal {
                part_name,
                content_type,
                data,
            } => {
                if part_name.as_str() == "/" {
                    return invalid("snapshot image cannot target the package root".into());
                }
                if !content_type.starts_with("image/") {
                    return invalid(format!(
                        "snapshot resource '{}' has non-image content type '{}'",
                        resource.relationship_id, content_type
                    ));
                }
                if data.len() > MAX_WEB_EXTENSION_SNAPSHOT_BYTES {
                    return invalid(format!(
                        "snapshot resource '{}' exceeds {MAX_WEB_EXTENSION_SNAPSHOT_BYTES} bytes",
                        resource.relationship_id
                    ));
                }
            },
            WebExtensionSnapshotTarget::External { target } => {
                if !*linked {
                    return invalid(format!(
                        "embedded snapshot resource '{}' cannot be external",
                        resource.relationship_id
                    ));
                }
                require_nonempty("external snapshot target", target)?;
            },
        }
    }
    Ok(())
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

fn canonical_node_xml(node: &Node) -> String {
    fn write_node(out: &mut String, node: &Node) {
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
    raw_fragment: Option<RawFragment>,
}

#[derive(Debug)]
struct RawFragment {
    start: usize,
    start_tag_end: usize,
    end: usize,
    namespaces: Vec<(String, String)>,
    declared_prefixes: HashSet<String>,
}

#[derive(Debug)]
struct NodeFrame {
    node: Node,
    namespaces: HashMap<String, String>,
    extension_depth: Option<usize>,
    direct_extension_count: usize,
}

#[derive(Debug, Default)]
struct XmlBuildState {
    root: Option<Node>,
    stack: Vec<NodeFrame>,
    string_bytes: usize,
    nodes: usize,
}

#[derive(Debug)]
struct XmlDocument {
    root: Option<Node>,
    xml: Vec<u8>,
}

impl XmlDocument {
    fn root(&self) -> Result<&Node> {
        self.root
            .as_ref()
            .ok_or_else(|| OoxmlError::InvalidFormat("missing XML root".into()))
    }

    fn self_contained_fragment(&self, node: &Node) -> Result<String> {
        let fragment = node.raw_fragment.as_ref().ok_or_else(|| {
            OoxmlError::InvalidFormat("XML node has no retained fragment bounds".into())
        })?;
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
            .map_err(|error| OoxmlError::Xml(format!("non-UTF-8 extension fragment: {error}")))?;
        let mut out = String::with_capacity(raw.len() + fragment.namespaces.len() * 48);
        out.push_str(&raw[..insert_at]);
        for (prefix, namespace) in &fragment.namespaces {
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
    parse_xml_owned(processed.xml.into_owned())
}

fn parse_xml(xml: &[u8]) -> Result<XmlDocument> {
    parse_xml_owned(xml.to_vec())
}

fn parse_xml_owned(xml: Vec<u8>) -> Result<XmlDocument> {
    let mut reader = Reader::from_reader(xml.as_slice());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut state = XmlBuildState::default();
    let mut xml_version = XmlVersion::Implicit1_0;
    loop {
        let event_start = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buffer)?;
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Decl(declaration) => xml_version = declaration.xml_version()?,
            Event::Start(element) => push_element(
                &reader,
                &element,
                &mut state,
                xml_version,
                false,
                event_start,
                event_end,
            )?,
            Event::Empty(element) => push_element(
                &reader,
                &element,
                &mut state,
                xml_version,
                true,
                event_start,
                event_end,
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
                let mut frame = state.stack.pop().expect("stack checked above");
                if let Some(fragment) = frame.node.raw_fragment.as_mut() {
                    fragment.end = event_end;
                }
                attach_node(&mut state.root, &mut state.stack, frame.node)?;
            },
            _ => {},
        }
        buffer.clear();
    }
    if !state.stack.is_empty() {
        return invalid("unclosed XML element".into());
    }
    if state.string_bytes > MAX_WEB_EXTENSION_STRING_BYTES {
        return invalid("web extension decoded strings exceed allocation limit".into());
    }
    drop(reader);
    Ok(XmlDocument {
        root: state.root,
        xml,
    })
}

fn push_element(
    reader: &Reader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    state: &mut XmlBuildState,
    xml_version: XmlVersion,
    empty: bool,
    event_start: usize,
    event_end: usize,
) -> Result<()> {
    if state.stack.len() >= MAX_WEB_EXTENSION_XML_DEPTH {
        return invalid(format!(
            "web extension XML depth exceeds {MAX_WEB_EXTENSION_XML_DEPTH}"
        ));
    }
    state.nodes = state
        .nodes
        .checked_add(1)
        .ok_or_else(|| OoxmlError::InvalidFormat("web extension node count overflow".into()))?;
    if state.nodes > MAX_WEB_EXTENSION_XML_NODES {
        return invalid(format!(
            "web extension XML node count exceeds {MAX_WEB_EXTENSION_XML_NODES}"
        ));
    }
    let mut namespaces = state
        .stack
        .last()
        .map(|frame| frame.namespaces.clone())
        .unwrap_or_default();
    namespaces.insert("xml".into(), "http://www.w3.org/XML/1998/namespace".into());
    let mut raw_attributes = Vec::new();
    let mut declared_prefixes = HashSet::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(xml_version, reader.decoder())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        state.string_bytes = state.string_bytes.saturating_add(name.len() + value.len());
        if state.string_bytes > MAX_WEB_EXTENSION_STRING_BYTES {
            return invalid("web extension decoded strings exceed allocation limit".into());
        }
        if name == "xmlns" {
            declared_prefixes.insert(String::new());
            namespaces.insert(String::new(), value);
        } else if let Some(prefix) = name.strip_prefix("xmlns:") {
            declared_prefixes.insert(prefix.to_owned());
            namespaces.insert(prefix.to_owned(), value);
        } else {
            raw_attributes.push((name, value));
        }
    }
    if namespaces.len() > 4096 {
        return invalid("web extension XML namespace bindings exceed 4096".into());
    }
    let element_name = element.name();
    let raw_name = std::str::from_utf8(element_name.as_ref())
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    let (prefix, local_name) = split_qname(raw_name);
    let namespace = namespaces.get(prefix).cloned().ok_or_else(|| {
        OoxmlError::InvalidFormat(format!("unbound XML namespace prefix '{prefix}'"))
    })?;
    state.string_bytes = state
        .string_bytes
        .saturating_add(namespace.len() + local_name.len());
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
    let capture_fragment = should_capture_extension_list(
        state.stack.last().map(|frame| &frame.node),
        &namespace,
        local_name,
    );
    let raw_fragment = capture_fragment.then(|| {
        let mut namespaces = namespaces
            .iter()
            .map(|(prefix, namespace)| (prefix.clone(), namespace.clone()))
            .collect::<Vec<_>>();
        namespaces.sort_unstable_by(|left, right| left.0.cmp(&right.0));
        RawFragment {
            start: event_start,
            start_tag_end: event_end,
            end: if empty { event_end } else { 0 },
            namespaces,
            declared_prefixes,
        }
    });
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
        let parent = state.stack.last_mut().expect("parent checked above");
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
            .ok_or_else(|| OoxmlError::InvalidFormat("extLst count overflow".into()))?;
        enforce_count("OfficeArt extension", parent.direct_extension_count)?;
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
    if empty {
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

fn attach_node(root: &mut Option<Node>, stack: &mut [NodeFrame], node: Node) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        if parent.extension_depth.is_none() {
            parent.node.children.push(node);
        }
    } else if root.replace(node).is_some() {
        return invalid("multiple XML root elements".into());
    }
    Ok(())
}

fn should_capture_extension_list(parent: Option<&Node>, namespace: &str, local_name: &str) -> bool {
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

fn extension_text_is_allowed(stack: &[NodeFrame]) -> bool {
    stack
        .last()
        .and_then(|frame| frame.extension_depth)
        .is_some_and(|depth| depth >= 2)
}

fn split_qname(name: &str) -> (&str, &str) {
    name.split_once(':').unwrap_or(("", name))
}

fn element_children(node: &Node) -> Vec<&Node> {
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
        .find(|attribute| attribute.namespace == namespace && attribute.local_name == local_name)
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

fn is_drawingml_namespace(namespace: &str) -> bool {
    matches!(namespace, DRAWINGML_NAMESPACE | STRICT_DRAWINGML_NAMESPACE)
}

fn optional_bool_attr(node: &Node, namespace: &str, local_name: &str) -> Result<Option<bool>> {
    attr(node, namespace, local_name)
        .map(parse_bool)
        .transpose()
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
    children
        .get(position)
        .is_some_and(|child| child.namespace == namespace && child.local_name == local_name)
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

    const LOCAL_OMEX_EXTENSION: &[u8] =
        include_bytes!("../../../test-data/ooxml/web_extensions/omex_webextension.xml");
    const LOCAL_REGISTRY_EXTENSION: &[u8] =
        include_bytes!("../../../test-data/ooxml/web_extensions/registry_webextension.xml");
    const LOCAL_VISIBLE_TASK_PANES: &[u8] =
        include_bytes!("../../../test-data/ooxml/web_extensions/visible_taskpanes.xml");
    const LOCAL_HIDDEN_TASK_PANES: &[u8] =
        include_bytes!("../../../test-data/ooxml/web_extensions/hidden_taskpanes.xml");
    const LOCAL_SNAPSHOT_EFFECTS_EXTENSION: &[u8] =
        include_bytes!("../../../test-data/ooxml/web_extensions/snapshot_effects_webextension.xml");
    const LOCAL_EXTENSION_LISTS_EXTENSION: &[u8] =
        include_bytes!("../../../test-data/ooxml/web_extensions/extension_lists_webextension.xml");
    const LOCAL_EXTENSION_LISTS_TASK_PANES: &[u8] =
        include_bytes!("../../../test-data/ooxml/web_extensions/extension_lists_taskpanes.xml");

    #[test]
    fn loads_local_omex_and_registry_fixtures_inertly() {
        let omex = local_fixture_package(LOCAL_VISIBLE_TASK_PANES, LOCAL_OMEX_EXTENSION);
        let panes = load_web_extension_task_panes(&omex).unwrap().unwrap();
        assert_eq!(panes.panes.len(), 1);
        assert_eq!(
            panes.panes[0].web_extension.reference.store_type,
            WebExtensionStoreType::Omex
        );
        assert!(panes.panes[0].visible);

        let registry = local_fixture_package(LOCAL_HIDDEN_TASK_PANES, LOCAL_REGISTRY_EXTENSION);
        let panes = load_web_extension_task_panes(&registry).unwrap().unwrap();
        assert_eq!(
            panes.panes[0].web_extension.reference.store_type,
            WebExtensionStoreType::Registry
        );
        assert!(!panes.panes[0].visible);
    }

    #[test]
    fn strict_writer_is_deterministic_and_round_trips() {
        let extension = sample_extension();
        let first = write_web_extension(&extension, OoxmlConformance::Strict).unwrap();
        let second = write_web_extension(&extension, OoxmlConformance::Strict).unwrap();
        assert_eq!(first, second);
        assert!(
            std::str::from_utf8(&first)
                .unwrap()
                .contains(STRICT_RELATIONSHIPS_NAMESPACE)
        );
        assert_eq!(parse_web_extension(&first).unwrap(), extension);
    }

    #[test]
    fn snapshot_compression_and_effect_trees_round_trip() {
        let extension = parse_web_extension(LOCAL_SNAPSHOT_EFFECTS_EXTENSION).unwrap();
        let snapshot = extension.snapshot.as_ref().unwrap();
        assert_eq!(
            snapshot.compression_state,
            Some(WebExtensionBlipCompression::HighQualityPrint)
        );
        assert_eq!(
            snapshot
                .effects
                .iter()
                .map(WebExtensionSnapshotEffect::kind)
                .collect::<Vec<_>>(),
            vec![
                WebExtensionSnapshotEffectKind::AlphaModulateFixed,
                WebExtensionSnapshotEffectKind::Duotone,
                WebExtensionSnapshotEffectKind::Blur,
            ]
        );
        assert!(snapshot.effects[1].xml().contains("srgbClr"));

        let written = write_web_extension(&extension, OoxmlConformance::Strict).unwrap();
        let reparsed = parse_web_extension(&written).unwrap();
        assert_eq!(reparsed, extension);
        let written = std::str::from_utf8(&written).unwrap();
        assert!(written.contains("cstate=\"hqprint\""));
        assert!(written.contains(STRICT_RELATIONSHIPS_NAMESPACE));
    }

    #[test]
    fn preserves_all_extension_list_sites_with_inherited_namespaces_and_mixed_content() {
        let extension = parse_web_extension(LOCAL_EXTENSION_LISTS_EXTENSION).unwrap();
        let reference_extension = extension.reference.extension_list.as_ref().unwrap();
        assert_eq!(
            reference_extension.kind(),
            WebExtensionExtensionListKind::WebExtension
        );
        assert!(reference_extension.xml().contains("xmlns:vendor="));
        assert!(reference_extension.xml().contains("xmlns:r="));
        assert!(reference_extension.xml().contains("reference text"));
        assert!(reference_extension.xml().contains("<![CDATA[<opaque>]]>"));
        assert!(reference_extension.xml().contains("<!--kept-->"));
        assert!(extension.alternate_references[0].extension_list.is_some());
        assert!(extension.bindings[0].extension_list.is_some());
        assert_eq!(
            extension
                .snapshot
                .as_ref()
                .unwrap()
                .extension_list
                .as_ref()
                .unwrap()
                .kind(),
            WebExtensionExtensionListKind::DrawingMl
        );
        assert!(extension.extension_list.is_some());

        let written = write_web_extension(&extension, OoxmlConformance::Strict).unwrap();
        assert_eq!(parse_web_extension(&written).unwrap(), extension);

        let panes = parse_task_panes(LOCAL_EXTENSION_LISTS_TASK_PANES).unwrap();
        let pane_extension = panes[0].extension_list.as_ref().unwrap();
        assert_eq!(
            pane_extension.kind(),
            WebExtensionExtensionListKind::TaskPane
        );
        assert!(pane_extension.xml().contains("xmlns:vendor="));
        assert!(pane_extension.xml().contains("<![CDATA[<pane-data>]]>"));
        assert!(pane_extension.xml().contains("<!--pane comment-->"));
    }

    #[test]
    fn package_crud_round_trips_every_inert_extension_list() {
        let package = local_fixture_package(
            LOCAL_EXTENSION_LISTS_TASK_PANES,
            LOCAL_EXTENSION_LISTS_EXTENSION,
        );
        let loaded = load_web_extension_task_panes(&package).unwrap().unwrap();
        assert!(loaded.panes[0].extension_list.is_some());
        assert!(loaded.panes[0].web_extension.extension_list.is_some());

        let mut stored = OpcPackage::new();
        store_web_extension_task_panes(&mut stored, &loaded, OoxmlConformance::Strict).unwrap();
        assert_eq!(
            load_web_extension_task_panes(&stored).unwrap(),
            Some(loaded)
        );
    }

    #[test]
    fn authored_extension_lists_validate_namespace_placement_and_security() {
        let web = WebExtensionExtensionList::from_xml(
            br#"<we:extLst xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"><a:ext xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uri="urn:test"><v:data xmlns:v="urn:test">text<![CDATA[data]]></v:data></a:ext></we:extLst>"#,
        )
        .unwrap();
        assert_eq!(web.kind(), WebExtensionExtensionListKind::WebExtension);
        assert_eq!(
            WebExtensionExtensionList::from_xml(web.as_xml()).unwrap(),
            web
        );

        let mut extension = sample_extension();
        extension.extension_list = Some(web.clone());
        assert!(write_web_extension(&extension, OoxmlConformance::Transitional).is_ok());

        let mut panes = sample_task_panes();
        panes.panes[0].extension_list = Some(web);
        assert!(write_task_panes(&panes, OoxmlConformance::Transitional).is_err());
        assert!(
            WebExtensionExtensionList::from_xml(
                br#"<!DOCTYPE extLst [<!ENTITY x "boom">]><we:extLst xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11">&x;</we:extLst>"#
            )
            .is_err()
        );
        assert!(
            WebExtensionExtensionList::from_xml(
                br#"<v:extLst xmlns:v="urn:not-an-office-namespace"/>"#
            )
            .is_err()
        );
        assert!(
            WebExtensionExtensionList::from_xml(
                br#"<we:extLst xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" unexpected="1"/>"#
            )
            .is_err()
        );
        assert!(
            WebExtensionExtensionList::from_xml(
                br#"<we:extLst xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ext/></we:extLst>"#
            )
            .is_err()
        );
        assert!(
            WebExtensionExtensionList::from_xml(
                br#"<we:extLst xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" xmlns:v="urn:test"><v:data/></we:extLst>"#
            )
            .is_err()
        );
        assert!(
            WebExtensionExtensionList::from_xml(
                br#"<we:extLst xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11">text</we:extLst>"#
            )
            .is_err()
        );
    }

    #[test]
    fn accepts_every_ct_blip_effect_kind_and_rejects_invalid_markup() {
        let names = [
            "alphaBiLevel",
            "alphaCeiling",
            "alphaFloor",
            "alphaInv",
            "alphaMod",
            "alphaModFix",
            "alphaRepl",
            "biLevel",
            "blur",
            "clrChange",
            "clrRepl",
            "duotone",
            "fillOverlay",
            "grayscl",
            "hsl",
            "lum",
            "tint",
        ];
        for name in names {
            let xml = format!(r#"<a:{name} xmlns:a="{DRAWINGML_NAMESPACE}"/>"#);
            let effect = WebExtensionSnapshotEffect::from_xml(xml.as_bytes()).unwrap();
            assert_eq!(effect.kind().local_name(), name);
            assert_eq!(
                WebExtensionSnapshotEffect::from_xml(effect.xml().as_bytes()).unwrap(),
                effect
            );
        }

        assert!(
            WebExtensionSnapshotEffect::from_xml(
                br#"<a:reflection xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#
            )
            .is_err()
        );
        assert!(
            WebExtensionSnapshotEffect::from_xml(
                br#"<!DOCTYPE x><a:blur xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#
            )
            .is_err()
        );
        assert!(
            WebExtensionSnapshotEffect::from_xml(
                br#"<a:blur xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">text</a:blur>"#
            )
            .is_err()
        );
    }

    #[test]
    fn rejects_invalid_snapshot_compression_and_effect_order() {
        let invalid_compression = format!(
            r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" id="x"><we:reference id="a" version="1"/><we:properties/><we:bindings/><we:snapshot cstate="lossless"/></we:webextension>"#
        );
        assert!(parse_web_extension(invalid_compression.as_bytes()).is_err());

        let misplaced_extension_list = format!(
            r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" xmlns:a="{DRAWINGML_NAMESPACE}" id="x"><we:reference id="a" version="1"/><we:properties/><we:bindings/><we:snapshot><a:extLst/><a:blur/></we:snapshot></we:webextension>"#
        );
        assert!(parse_web_extension(misplaced_extension_list.as_bytes()).is_err());
    }

    #[test]
    fn accepts_mce_alternate_content_and_strict_relationship_attributes() {
        let xml = format!(
            r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" xmlns:r="{STRICT_RELATIONSHIPS_NAMESPACE}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" id="x"><we:reference id="a" version="1"/><mc:AlternateContent><mc:Choice Requires="we"><we:alternateReferences/></mc:Choice><mc:Fallback/></mc:AlternateContent><we:properties/><we:bindings/><we:snapshot r:embed="rId1"/></we:webextension>"#
        );
        let extension = parse_web_extension(xml.as_bytes()).unwrap();
        assert_eq!(
            extension
                .snapshot
                .unwrap()
                .embedded_relationship_id
                .as_deref(),
            Some("rId1")
        );
    }

    #[test]
    fn rejects_dtd_bad_order_bad_store_and_nonfinite_width() {
        assert!(parse_web_extension(br#"<!DOCTYPE x><x/>"#).is_err());
        let bad_order = format!(
            r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" id="x"><we:properties/><we:reference id="a" version="1"/><we:bindings/></we:webextension>"#
        );
        assert!(parse_web_extension(bad_order.as_bytes()).is_err());
        let bad_store = format!(
            r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" id="x"><we:reference id="a" version="1" storeType="Network"/><we:properties/><we:bindings/></we:webextension>"#
        );
        assert!(parse_web_extension(bad_store.as_bytes()).is_err());
        let bad_width = format!(
            r#"<wetp:taskpanes xmlns:wetp="{TASK_PANES_NAMESPACE}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS_NAMESPACE}"><wetp:taskpane dockstate="right" visibility="1" width="NaN" row="0"><wetp:webextensionref r:id="rId1"/></wetp:taskpane></wetp:taskpanes>"#
        );
        assert!(parse_task_panes(bad_width.as_bytes()).is_err());
        let obsolete_float = format!(
            r#"<wetp:taskpanes xmlns:wetp="{TASK_PANES_NAMESPACE}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS_NAMESPACE}"><wetp:taskpane dockstate="right" visibility="1" width="320" row="0"><wetp:webextensionref r:id="rId1"/><wetp:float/></wetp:taskpane></wetp:taskpanes>"#
        );
        assert!(parse_task_panes(obsolete_float.as_bytes()).is_err());
    }

    #[test]
    fn enforces_input_and_list_caps() {
        assert!(parse_web_extension(&vec![b' '; MAX_WEB_EXTENSION_XML_BYTES + 1]).is_err());
        let mut model = WebExtensionTaskPanes::default();
        model
            .panes
            .resize_with(MAX_WEB_EXTENSION_ITEMS + 1, || WebExtensionTaskPane {
                dock_state: "right".into(),
                visible: false,
                width: 320.0,
                row: 0,
                locked: false,
                relationship_id: "rId1".into(),
                web_extension: sample_extension(),
                snapshot_resources: vec![],
                extension_list: None,
            });
        assert!(write_task_panes(&model, OoxmlConformance::Transitional).is_err());
        let mut extension = sample_extension();
        extension.id = "x".repeat(MAX_WEB_EXTENSION_XML_BYTES);
        assert!(write_web_extension(&extension, OoxmlConformance::Transitional).is_err());

        let mut excessive_nodes = format!(
            r#"<we:extLst xmlns:we="{WEB_EXTENSION_NAMESPACE}" xmlns:a="{DRAWINGML_NAMESPACE}" xmlns:v="urn:test"><a:ext uri="urn:test"><v:data>"#
        );
        excessive_nodes.push_str(&"<v:n/>".repeat(MAX_WEB_EXTENSION_XML_NODES));
        excessive_nodes.push_str("</v:data></a:ext></we:extLst>");
        assert!(WebExtensionExtensionList::from_xml(excessive_nodes.as_bytes()).is_err());
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

    #[test]
    fn package_crud_round_trips_embedded_and_linked_snapshots() {
        let mut package = OpcPackage::new();
        let authored = sample_task_panes();
        store_web_extension_task_panes(&mut package, &authored, OoxmlConformance::Transitional)
            .unwrap();
        assert_eq!(
            load_web_extension_task_panes(&package).unwrap(),
            Some(authored.clone())
        );

        let task_panes_name = package
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP_TYPE)
            .unwrap()
            .target_partname()
            .unwrap();
        let extension_name = package
            .get_part(&task_panes_name)
            .unwrap()
            .rels()
            .get("rId1")
            .unwrap()
            .target_partname()
            .unwrap();
        let extension = package.get_part(&extension_name).unwrap();
        assert!(!extension.rels().get("rIdSnapshot").unwrap().is_external());
        assert!(extension.rels().get("rIdLinked").unwrap().is_external());

        let mut replacement = authored;
        replacement.panes[0].web_extension.snapshot = None;
        replacement.panes[0].snapshot_resources.clear();
        replacement.panes[0].visible = false;
        store_web_extension_task_panes(&mut package, &replacement, OoxmlConformance::Strict)
            .unwrap();
        assert_eq!(
            load_web_extension_task_panes(&package).unwrap(),
            Some(replacement)
        );
        assert!(
            package
                .get_part(&PackURI::new("/media/web-extension-snapshot.png").unwrap())
                .is_err()
        );

        assert!(remove_web_extension_task_panes(&mut package).unwrap());
        assert!(load_web_extension_task_panes(&package).unwrap().is_none());
        assert!(!remove_web_extension_task_panes(&mut package).unwrap());
        assert_eq!(package.part_count(), 0);
    }

    #[test]
    fn package_store_rejects_resource_mismatches_without_mutation() {
        let mut package = OpcPackage::new();
        let mut malformed = sample_task_panes();
        malformed.panes[0].snapshot_resources.pop();
        assert!(
            store_web_extension_task_panes(
                &mut package,
                &malformed,
                OoxmlConformance::Transitional,
            )
            .is_err()
        );
        assert_eq!(package.part_count(), 0);
        assert_eq!(package.rels().iter().count(), 0);

        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/media/web-extension-snapshot.png").unwrap(),
            "image/png".into(),
            vec![9, 9, 9],
        )));
        assert!(
            store_web_extension_task_panes(
                &mut package,
                &sample_task_panes(),
                OoxmlConformance::Transitional,
            )
            .is_err()
        );
        assert_eq!(package.part_count(), 1);
        assert_eq!(
            package
                .get_part(&PackURI::new("/media/web-extension-snapshot.png").unwrap())
                .unwrap()
                .blob(),
            &[9, 9, 9]
        );
    }

    fn synthetic_package(
        external_extension: bool,
        extension_content_type: &str,
        pane_relationship_id: &str,
    ) -> OpcPackage {
        let task_panes_xml = format!(
            r#"<wetp:taskpanes xmlns:wetp="{TASK_PANES_NAMESPACE}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS_NAMESPACE}"><wetp:taskpane dockstate="right" visibility="0" width="320" row="0"><wetp:webextensionref r:id="{pane_relationship_id}"/></wetp:taskpane></wetp:taskpanes>"#
        );
        let extension_xml =
            write_web_extension(&sample_extension(), OoxmlConformance::Transitional).unwrap();
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

    fn local_fixture_package(task_panes_xml: &[u8], extension_xml: &[u8]) -> OpcPackage {
        let mut package = OpcPackage::new();
        package.rels_mut().add_relationship(
            TASK_PANES_RELATIONSHIP_TYPE.into(),
            "webextensions/taskpanes.xml".into(),
            "rIdTaskPanes".into(),
            false,
        );
        let mut task_panes_part = XmlPart::new(
            PackURI::new("/webextensions/taskpanes.xml").unwrap(),
            TASK_PANES_CONTENT_TYPE.into(),
            task_panes_xml.to_vec(),
        );
        task_panes_part.rels_mut().add_relationship(
            WEB_EXTENSION_RELATIONSHIP_TYPE.into(),
            "webextension1.xml".into(),
            "rId1".into(),
            false,
        );
        package.add_part(Box::new(task_panes_part));
        package.add_part(Box::new(XmlPart::new(
            PackURI::new("/webextensions/webextension1.xml").unwrap(),
            WEB_EXTENSION_CONTENT_TYPE.into(),
            extension_xml.to_vec(),
        )));
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
                extension_list: None,
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
                extension_list: None,
            }],
            snapshot: Some(WebExtensionSnapshot::default()),
            extension_list: None,
        }
    }

    fn sample_task_panes() -> WebExtensionTaskPanes {
        let mut extension = sample_extension();
        extension.snapshot = Some(WebExtensionSnapshot {
            embedded_relationship_id: Some("rIdSnapshot".into()),
            linked_relationship_id: Some("rIdLinked".into()),
            compression_state: Some(WebExtensionBlipCompression::HighQualityPrint),
            effects: vec![
                WebExtensionSnapshotEffect::from_xml(
                    br#"<a:alphaModFix xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" amt="50000"/>"#,
                )
                .unwrap(),
            ],
            extension_list: None,
        });
        WebExtensionTaskPanes {
            panes: vec![WebExtensionTaskPane {
                dock_state: "right".into(),
                visible: true,
                width: 320.0,
                row: 0,
                locked: false,
                relationship_id: "rId1".into(),
                web_extension: extension,
                snapshot_resources: vec![
                    WebExtensionSnapshotResource {
                        relationship_id: "rIdSnapshot".into(),
                        target: WebExtensionSnapshotTarget::Internal {
                            part_name: PackURI::new("/media/web-extension-snapshot.png").unwrap(),
                            content_type: "image/png".into(),
                            data: vec![0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a],
                        },
                    },
                    WebExtensionSnapshotResource {
                        relationship_id: "rIdLinked".into(),
                        target: WebExtensionSnapshotTarget::External {
                            target: "https://example.invalid/inert-snapshot.png".into(),
                        },
                    },
                ],
                extension_list: None,
            }],
        }
    }
}
