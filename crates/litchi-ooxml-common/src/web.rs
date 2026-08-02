//! Inert Office Web Extension and persisted task-pane metadata.
//!
//! This module implements the package structures defined by MS-OWEXML. It
//! intentionally does not locate add-ins, contact catalog providers, load
//! manifests, resolve linked content, or execute scripts/custom functions.

use crate::{Error, MceCapabilities, MceLimits, Result, process_markup_compatibility};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use quick_xml::Reader;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use std::collections::{BTreeSet, HashMap, HashSet, VecDeque};
use std::sync::Arc;

const WEB_EXTENSION_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/webextensions/webextension/2010/11";
const TASK_PANES_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11";
const TRANSITIONAL_RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships";
const DRAWINGML_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML_NAMESPACE: &str = "http://purl.oclc.org/ooxml/drawingml/main";

/// Low-level OPC constants for callers constructing synthetic or specialized graphs.
pub mod raw {
    /// Package relationship to the persisted task-pane part.
    pub const TASK_PANES_RELATIONSHIP: &str =
        "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes";
    /// Relationship from the task-pane part to one Office Add-in part.
    pub const ADD_IN_RELATIONSHIP: &str =
        "http://schemas.microsoft.com/office/2011/relationships/webextension";
    /// Content type of a persisted task-pane part.
    pub const TASK_PANES_CONTENT_TYPE: &str = "application/vnd.ms-office.webextensiontaskpanes+xml";
    /// Content type of an Office Add-in part.
    pub const ADD_IN_CONTENT_TYPE: &str = "application/vnd.ms-office.webextension+xml";
}

use raw::{
    ADD_IN_CONTENT_TYPE, ADD_IN_RELATIONSHIP, TASK_PANES_CONTENT_TYPE, TASK_PANES_RELATIONSHIP,
};

const IMAGE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const STRICT_IMAGE_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image";

const STANDARD_XML_BYTES: usize = 4 * 1024 * 1024;
const STANDARD_TOTAL_XML_BYTES: usize = 64 * 1024 * 1024;
const STANDARD_DEPTH: usize = 128;
const STANDARD_NODES: usize = 65_536;
const STANDARD_ITEMS: usize = 4096;
const STANDARD_STRING_BYTES: usize = 8 * 1024 * 1024;
const STANDARD_TOTAL_STRING_BYTES: usize = 128 * 1024 * 1024;
const STANDARD_IMAGE_BYTES: usize = 64 * 1024 * 1024;
const STANDARD_TOTAL_IMAGE_BYTES: usize = 256 * 1024 * 1024;
const STANDARD_PACKAGE_PARTS: usize = 65_536;
const STANDARD_PACKAGE_RELATIONSHIPS: usize = 262_144;
const STANDARD_PART_ALLOCATIONS: usize = 8_192;
const STANDARD_PART_DELETIONS: usize = 8_192;

/// Resource ceilings for inert Office Add-in metadata and snapshot payloads.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum bytes in one source or authored XML part.
    pub xml_bytes: usize,
    /// Maximum aggregate source or authored XML bytes in one operation.
    pub total_xml_bytes: usize,
    /// Maximum XML element nesting depth.
    pub depth: usize,
    /// Maximum element count in one XML part or retained fragment.
    pub nodes: usize,
    /// Maximum panes or items in any schema collection.
    pub items: usize,
    /// Maximum aggregate decoded string bytes in one XML part.
    pub string_bytes: usize,
    /// Maximum aggregate retained XML, decoded strings, and indexed package metadata.
    pub total_string_bytes: usize,
    /// Maximum bytes in one embedded snapshot image.
    pub image_bytes: usize,
    /// Maximum unique embedded snapshot bytes in one package graph.
    pub total_image_bytes: usize,
    /// Maximum number of package parts inspected by one operation.
    pub package_parts: usize,
    /// Maximum aggregate package-level and part-level relationships inspected.
    pub package_relationships: usize,
    /// Maximum new parts or deterministic part-name allocation attempts.
    pub part_allocations: usize,
    /// Maximum old graph parts that one operation may delete.
    pub part_deletions: usize,
}

impl Limits {
    /// Conservative defaults for untrusted packages.
    #[must_use]
    pub const fn standard() -> Self {
        Self {
            xml_bytes: STANDARD_XML_BYTES,
            total_xml_bytes: STANDARD_TOTAL_XML_BYTES,
            depth: STANDARD_DEPTH,
            nodes: STANDARD_NODES,
            items: STANDARD_ITEMS,
            string_bytes: STANDARD_STRING_BYTES,
            total_string_bytes: STANDARD_TOTAL_STRING_BYTES,
            image_bytes: STANDARD_IMAGE_BYTES,
            total_image_bytes: STANDARD_TOTAL_IMAGE_BYTES,
            package_parts: STANDARD_PACKAGE_PARTS,
            package_relationships: STANDARD_PACKAGE_RELATIONSHIPS,
            part_allocations: STANDARD_PART_ALLOCATIONS,
            part_deletions: STANDARD_PART_DELETIONS,
        }
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::standard()
    }
}

/// Task-pane docking state with forward-compatible retention of newer values.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Dock {
    Left,
    Right,
    Top,
    Bottom,
    Floating,
    Other(String),
}

impl Dock {
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::Left => "left",
            Self::Right => "right",
            Self::Top => "top",
            Self::Bottom => "bottom",
            Self::Floating => "float",
            Self::Other(value) => value,
        }
    }

    fn parse(value: &str) -> Result<Self> {
        require_nonempty("dock state", value)?;
        Ok(match value {
            "left" => Self::Left,
            "right" => Self::Right,
            "top" => Self::Top,
            "bottom" => Self::Bottom,
            "float" | "floating" => Self::Floating,
            value => Self::Other(value.to_owned()),
        })
    }
}

impl AsRef<str> for Dock {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

// Internal defaults used by standalone fragment constructors. Package-level
// entry points thread caller-provided limits explicitly.
const MAX_WEB_EXTENSION_XML_BYTES: usize = STANDARD_XML_BYTES;
#[cfg(test)]
const MAX_WEB_EXTENSION_XML_NODES: usize = STANDARD_NODES;
const MAX_WEB_EXTENSION_ITEMS: usize = STANDARD_ITEMS;
const MAX_WEB_EXTENSION_SNAPSHOT_BYTES: usize = STANDARD_IMAGE_BYTES;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
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
pub enum Store {
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

impl Store {
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
pub struct Reference {
    id: String,
    version: String,
    catalog: Option<String>,
    store: Store,
    extension_list: Option<ExtList>,
}

impl Reference {
    /// Create a validated catalog reference.
    pub fn new(id: impl Into<String>, version: impl Into<String>, store: Store) -> Result<Self> {
        let value = Self {
            id: id.into(),
            version: version.into(),
            catalog: None,
            store,
            extension_list: None,
        };
        validate_store_reference(&value)?;
        Ok(value)
    }

    /// Add the optional catalog/location discriminator.
    pub fn catalog(mut self, value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        require_nonempty("reference catalog", &value)?;
        self.catalog = Some(value);
        Ok(self)
    }

    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    #[must_use]
    pub fn version(&self) -> &str {
        &self.version
    }

    #[must_use]
    pub const fn store(&self) -> Store {
        self.store
    }

    #[must_use]
    pub fn catalog_name(&self) -> Option<&str> {
        self.catalog.as_deref()
    }

    #[must_use]
    pub const fn ext(&self) -> Option<&ExtList> {
        self.extension_list.as_ref()
    }

    pub fn set_ext(&mut self, extension: ExtList) -> Result<&mut Self> {
        validate_extension_list(Some(&extension), &[ExtKind::AddIn])?;
        self.extension_list = Some(extension);
        Ok(self)
    }

    pub fn with_ext(mut self, extension: ExtList) -> Result<Self> {
        self.set_ext(extension)?;
        Ok(self)
    }

    pub fn clear_ext(&mut self) -> Option<ExtList> {
        self.extension_list.take()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Property {
    name: String,
    value: String,
}

impl Property {
    pub fn new(name: impl Into<String>, value: impl Into<String>) -> Result<Self> {
        let value = Self {
            name: name.into(),
            value: value.into(),
        };
        require_nonempty("property name", &value.name)?;
        Ok(value)
    }

    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// Binding data shape with forward-compatible retention of newer values.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BindingKind {
    Matrix,
    Table,
    Text,
    Other(String),
}

impl BindingKind {
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::Matrix => "matrix",
            Self::Table => "table",
            Self::Text => "text",
            Self::Other(value) => value,
        }
    }

    fn parse(value: &str) -> Result<Self> {
        require_nonempty("binding type", value)?;
        Ok(match value {
            "matrix" => Self::Matrix,
            "table" => Self::Table,
            "text" => Self::Text,
            value => Self::Other(value.to_owned()),
        })
    }
}

impl AsRef<str> for BindingKind {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Binding {
    id: String,
    kind: BindingKind,
    app_ref: String,
    extension_list: Option<ExtList>,
}

impl Binding {
    pub fn new(
        id: impl Into<String>,
        kind: impl AsRef<str>,
        app_ref: impl Into<String>,
    ) -> Result<Self> {
        let value = Self {
            id: id.into(),
            kind: BindingKind::parse(kind.as_ref())?,
            app_ref: app_ref.into(),
            extension_list: None,
        };
        validate_binding(&value)?;
        Ok(value)
    }

    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    #[must_use]
    pub const fn kind(&self) -> &BindingKind {
        &self.kind
    }

    #[must_use]
    pub fn kind_name(&self) -> &str {
        self.kind.as_str()
    }

    #[must_use]
    pub fn app_ref(&self) -> &str {
        &self.app_ref
    }

    #[must_use]
    pub const fn ext(&self) -> Option<&ExtList> {
        self.extension_list.as_ref()
    }

    pub fn set_ext(&mut self, extension: ExtList) -> Result<&mut Self> {
        validate_extension_list(Some(&extension), &[ExtKind::AddIn])?;
        self.extension_list = Some(extension);
        Ok(self)
    }

    pub fn with_ext(mut self, extension: ExtList) -> Result<Self> {
        self.set_ext(extension)?;
        Ok(self)
    }

    pub fn clear_ext(&mut self) -> Option<ExtList> {
        self.extension_list.take()
    }
}

/// Namespace dialect of an MS-OWEXML extension-list element.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExtKind {
    AddIn,
    TaskPane,
    DrawingMl,
    StrictDrawingMl,
}

impl ExtKind {
    pub fn namespace(self) -> &'static str {
        match self {
            Self::AddIn => WEB_EXTENSION_NAMESPACE,
            Self::TaskPane => TASK_PANES_NAMESPACE,
            Self::DrawingMl => DRAWINGML_NAMESPACE,
            Self::StrictDrawingMl => STRICT_DRAWINGML_NAMESPACE,
        }
    }

    fn from_namespace(namespace: &str) -> Result<Self> {
        match namespace {
            WEB_EXTENSION_NAMESPACE => Ok(Self::AddIn),
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
pub struct ExtList {
    kind: ExtKind,
    xml: String,
}

impl ExtList {
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_WEB_EXTENSION_XML_BYTES {
            return invalid(format!(
                "web extension extLst XML exceeds {MAX_WEB_EXTENSION_XML_BYTES} bytes"
            ));
        }
        let document = parse_xml(xml)?;
        Self::from_node(document.root()?, &document)
    }

    pub fn kind(&self) -> ExtKind {
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
        let kind = ExtKind::from_namespace(&node.namespace)?;
        Ok(Self {
            kind,
            xml: document.self_contained_fragment(node)?,
        })
    }
}

/// Compression state of a DrawingML `CT_Blip`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Compression {
    Email,
    Screen,
    Print,
    HighQualityPrint,
    None,
}

impl Compression {
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
pub enum EffectKind {
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

impl EffectKind {
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
pub struct Effect {
    kind: EffectKind,
    xml: String,
}

impl Effect {
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_WEB_EXTENSION_XML_BYTES {
            return invalid(format!(
                "snapshot effect XML exceeds {MAX_WEB_EXTENSION_XML_BYTES} bytes"
            ));
        }
        let document = parse_xml(xml)?;
        Self::from_node(document.root()?)
    }

    pub fn kind(&self) -> EffectKind {
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
        let kind = EffectKind::parse(&node.local_name)?;
        Ok(Self {
            kind,
            xml: canonical_node_xml(node),
        })
    }
}

/// DrawingML `CT_Blip` metadata used by a web-extension snapshot.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Snapshot {
    embedded_relationship_id: Option<String>,
    linked_relationship_id: Option<String>,
    compression_state: Option<Compression>,
    effects: Vec<Effect>,
    extension_list: Option<ExtList>,
}

impl Snapshot {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            embedded_relationship_id: None,
            linked_relationship_id: None,
            compression_state: None,
            effects: Vec::new(),
            extension_list: None,
        }
    }

    #[must_use]
    pub const fn compression(&self) -> Option<Compression> {
        self.compression_state
    }

    #[must_use]
    pub fn effects(&self) -> &[Effect] {
        &self.effects
    }

    pub fn set_compression(&mut self, compression: Option<Compression>) -> &mut Self {
        self.compression_state = compression;
        self
    }

    pub fn push_effect(&mut self, effect: Effect) -> Result<&mut Self> {
        if self.effects.len() >= MAX_WEB_EXTENSION_ITEMS {
            return limit(
                "snapshot effects",
                MAX_WEB_EXTENSION_ITEMS,
                self.effects.len().saturating_add(1),
            );
        }
        let reparsed = Effect::from_xml(effect.xml.as_bytes())?;
        if reparsed.kind != effect.kind {
            return invalid("snapshot effect kind does not match its XML root".into());
        }
        self.effects.push(effect);
        Ok(self)
    }

    pub fn replace_effect(&mut self, index: usize, effect: Effect) -> Result<Option<Effect>> {
        let reparsed = Effect::from_xml(effect.xml.as_bytes())?;
        if reparsed.kind != effect.kind {
            return invalid("snapshot effect kind does not match its XML root".into());
        }
        let Some(slot) = self.effects.get_mut(index) else {
            return Ok(None);
        };
        Ok(Some(std::mem::replace(slot, effect)))
    }

    pub fn remove_effect(&mut self, index: usize) -> Option<Effect> {
        (index < self.effects.len()).then(|| self.effects.remove(index))
    }

    pub fn clear_effects(&mut self) -> bool {
        let changed = !self.effects.is_empty();
        self.effects.clear();
        changed
    }

    #[must_use]
    pub const fn ext(&self) -> Option<&ExtList> {
        self.extension_list.as_ref()
    }

    pub fn set_ext(&mut self, extension: ExtList) -> Result<&mut Self> {
        validate_extension_list(
            Some(&extension),
            &[ExtKind::DrawingMl, ExtKind::StrictDrawingMl],
        )?;
        self.extension_list = Some(extension);
        Ok(self)
    }

    pub fn clear_ext(&mut self) -> Option<ExtList> {
        self.extension_list.take()
    }
}

/// One inert image relationship owned by a web-extension snapshot.
#[derive(Debug, Clone, PartialEq, Eq)]
struct SnapshotResource {
    relationship_id: String,
    target: SnapshotTarget,
}

/// Internal image bytes or an external linked image target.
///
/// External targets are retained as strings and are never fetched.
#[derive(Debug, Clone, PartialEq, Eq)]
enum SnapshotTarget {
    Internal {
        part_name: PackURI,
        content_type: String,
        data: Arc<Vec<u8>>,
    },
    External {
        target: String,
    },
}

/// A borrowed embedded snapshot. Cloning `shared` clones only the `Arc`.
#[derive(Debug, Clone, Copy)]
pub struct Image<'a> {
    part_name: &'a PackURI,
    content_type: &'a str,
    data: &'a Arc<Vec<u8>>,
}

impl Image<'_> {
    #[must_use]
    pub fn name(&self) -> &PackURI {
        self.part_name
    }

    #[must_use]
    pub fn content_type(&self) -> &str {
        self.content_type
    }

    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.data.as_slice()
    }

    /// Share the backing allocation without copying the image payload.
    #[must_use]
    pub fn shared(&self) -> Arc<Vec<u8>> {
        Arc::clone(self.data)
    }
}

/// A borrowed linked-image target. External targets remain inert and are never fetched.
#[derive(Debug, Clone, Copy)]
pub enum Link<'a> {
    Internal(Image<'a>),
    External(&'a str),
}

impl<'a> Link<'a> {
    #[must_use]
    pub const fn internal(self) -> Option<Image<'a>> {
        match self {
            Self::Internal(image) => Some(image),
            Self::External(_) => None,
        }
    }

    #[must_use]
    pub const fn external(self) -> Option<&'a str> {
        match self {
            Self::External(target) => Some(target),
            Self::Internal(_) => None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AddIn {
    id: String,
    frozen: bool,
    reference: Reference,
    alternate_references: Vec<Reference>,
    properties: Vec<Property>,
    bindings: Vec<Binding>,
    snapshot: Option<Snapshot>,
    extension_list: Option<ExtList>,
}

impl AddIn {
    pub fn new(id: impl Into<String>, reference: Reference) -> Result<Self> {
        let value = Self {
            id: id.into(),
            frozen: false,
            reference,
            alternate_references: Vec::new(),
            properties: Vec::new(),
            bindings: Vec::new(),
            snapshot: None,
            extension_list: None,
        };
        validate_model(&value)?;
        Ok(value)
    }

    #[must_use]
    pub fn frozen(mut self, frozen: bool) -> Self {
        self.frozen = frozen;
        self
    }

    pub fn set_frozen(&mut self, frozen: bool) -> &mut Self {
        self.frozen = frozen;
        self
    }

    pub fn bind(mut self, binding: Binding) -> Result<Self> {
        self.push_binding(binding)?;
        Ok(self)
    }

    pub fn push_binding(&mut self, binding: Binding) -> Result<&mut Self> {
        validate_binding(&binding)?;
        if self.bindings.len() >= MAX_WEB_EXTENSION_ITEMS {
            return limit(
                "web extension bindings",
                MAX_WEB_EXTENSION_ITEMS,
                self.bindings.len().saturating_add(1),
            );
        }
        if self.bindings.iter().any(|value| value.id == binding.id) {
            return invalid(format!("duplicate binding id '{}'", binding.id));
        }
        if self
            .bindings
            .iter()
            .any(|value| value.app_ref == binding.app_ref)
        {
            return invalid(format!("duplicate binding appRef '{}'", binding.app_ref));
        }
        self.bindings.push(binding);
        Ok(self)
    }

    pub fn upsert_binding(&mut self, binding: Binding) -> Result<&mut Self> {
        validate_binding(&binding)?;
        if let Some(index) = self
            .bindings
            .iter()
            .position(|value| value.id == binding.id)
        {
            if self
                .bindings
                .iter()
                .enumerate()
                .any(|(other, value)| other != index && value.app_ref == binding.app_ref)
            {
                return invalid(format!("duplicate binding appRef '{}'", binding.app_ref));
            }
            self.bindings[index] = binding;
            return Ok(self);
        }
        self.push_binding(binding)
    }

    #[must_use]
    pub fn binding<'key>(&self, selector: impl Into<Selector<'key>>) -> Option<&Binding> {
        match selector.into() {
            Selector::Id(id) => self.bindings.iter().find(|value| value.id == id),
            Selector::Index(index) => self.bindings.get(index),
        }
    }

    #[must_use]
    pub fn binding_mut<'key>(
        &mut self,
        selector: impl Into<Selector<'key>>,
    ) -> Option<&mut Binding> {
        match selector.into() {
            Selector::Id(id) => self.bindings.iter_mut().find(|value| value.id == id),
            Selector::Index(index) => self.bindings.get_mut(index),
        }
    }

    pub fn remove_binding<'key>(&mut self, selector: impl Into<Selector<'key>>) -> Option<Binding> {
        let index = match selector.into() {
            Selector::Id(id) => self.bindings.iter().position(|value| value.id == id)?,
            Selector::Index(index) if index < self.bindings.len() => index,
            Selector::Index(_) => return None,
        };
        Some(self.bindings.remove(index))
    }

    pub fn prop(mut self, property: Property) -> Result<Self> {
        self.push_property(property)?;
        Ok(self)
    }

    pub fn push_property(&mut self, property: Property) -> Result<&mut Self> {
        if self.properties.len() >= MAX_WEB_EXTENSION_ITEMS {
            return limit(
                "web extension properties",
                MAX_WEB_EXTENSION_ITEMS,
                self.properties.len().saturating_add(1),
            );
        }
        if self
            .properties
            .iter()
            .any(|value| value.name == property.name)
        {
            return invalid(format!("duplicate property name '{}'", property.name));
        }
        self.properties.push(property);
        Ok(self)
    }

    pub fn upsert_property(&mut self, property: Property) -> Result<&mut Self> {
        require_nonempty("property name", &property.name)?;
        if let Some(index) = self
            .properties
            .iter()
            .position(|value| value.name == property.name)
        {
            self.properties[index] = property;
            return Ok(self);
        }
        self.push_property(property)
    }

    #[must_use]
    pub fn property<'key>(&self, selector: impl Into<Selector<'key>>) -> Option<&Property> {
        match selector.into() {
            Selector::Id(name) => self.properties.iter().find(|value| value.name == name),
            Selector::Index(index) => self.properties.get(index),
        }
    }

    pub fn remove_property<'key>(
        &mut self,
        selector: impl Into<Selector<'key>>,
    ) -> Option<Property> {
        let index = match selector.into() {
            Selector::Id(name) => self
                .properties
                .iter()
                .position(|value| value.name == name)?,
            Selector::Index(index) if index < self.properties.len() => index,
            Selector::Index(_) => return None,
        };
        Some(self.properties.remove(index))
    }

    pub fn push_reference(&mut self, reference: Reference) -> Result<&mut Self> {
        validate_store_reference(&reference)?;
        if self.alternate_references.len() >= MAX_WEB_EXTENSION_ITEMS {
            return limit(
                "alternate references",
                MAX_WEB_EXTENSION_ITEMS,
                self.alternate_references.len().saturating_add(1),
            );
        }
        if reference.id == self.reference.id
            || self
                .alternate_references
                .iter()
                .any(|value| value.id == reference.id)
        {
            return invalid(format!("duplicate reference id '{}'", reference.id));
        }
        self.alternate_references.push(reference);
        Ok(self)
    }

    pub fn upsert_reference(&mut self, reference: Reference) -> Result<&mut Self> {
        validate_store_reference(&reference)?;
        if reference.id == self.reference.id {
            return invalid(format!(
                "alternate reference id '{}' duplicates the primary reference",
                reference.id
            ));
        }
        if let Some(index) = self
            .alternate_references
            .iter()
            .position(|value| value.id == reference.id)
        {
            self.alternate_references[index] = reference;
            return Ok(self);
        }
        self.push_reference(reference)
    }

    #[must_use]
    pub fn alternate_reference<'key>(
        &self,
        selector: impl Into<Selector<'key>>,
    ) -> Option<&Reference> {
        match selector.into() {
            Selector::Id(id) => self
                .alternate_references
                .iter()
                .find(|value| value.id == id),
            Selector::Index(index) => self.alternate_references.get(index),
        }
    }

    #[must_use]
    pub fn alternate_reference_mut<'key>(
        &mut self,
        selector: impl Into<Selector<'key>>,
    ) -> Option<&mut Reference> {
        match selector.into() {
            Selector::Id(id) => self
                .alternate_references
                .iter_mut()
                .find(|value| value.id == id),
            Selector::Index(index) => self.alternate_references.get_mut(index),
        }
    }

    pub fn remove_reference<'key>(
        &mut self,
        selector: impl Into<Selector<'key>>,
    ) -> Option<Reference> {
        let index = match selector.into() {
            Selector::Id(id) => self
                .alternate_references
                .iter()
                .position(|value| value.id == id)?,
            Selector::Index(index) if index < self.alternate_references.len() => index,
            Selector::Index(_) => return None,
        };
        Some(self.alternate_references.remove(index))
    }

    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    #[must_use]
    pub const fn is_frozen(&self) -> bool {
        self.frozen
    }

    #[must_use]
    pub const fn reference(&self) -> &Reference {
        &self.reference
    }

    pub fn set_reference(&mut self, reference: Reference) -> Result<&mut Self> {
        validate_store_reference(&reference)?;
        if self
            .alternate_references
            .iter()
            .any(|alternate| alternate.id == reference.id)
        {
            return invalid(format!(
                "primary reference id '{}' duplicates an alternate reference",
                reference.id
            ));
        }
        self.reference = reference;
        Ok(self)
    }

    pub const fn reference_mut(&mut self) -> &mut Reference {
        &mut self.reference
    }

    #[must_use]
    pub fn alternate_references(&self) -> &[Reference] {
        &self.alternate_references
    }

    #[must_use]
    pub fn properties(&self) -> &[Property] {
        &self.properties
    }

    #[must_use]
    pub fn bindings(&self) -> &[Binding] {
        &self.bindings
    }

    #[must_use]
    pub const fn snapshot(&self) -> Option<&Snapshot> {
        self.snapshot.as_ref()
    }

    #[must_use]
    pub const fn ext(&self) -> Option<&ExtList> {
        self.extension_list.as_ref()
    }

    pub fn set_ext(&mut self, extension: ExtList) -> Result<&mut Self> {
        validate_extension_list(Some(&extension), &[ExtKind::AddIn])?;
        self.extension_list = Some(extension);
        Ok(self)
    }

    pub fn clear_ext(&mut self) -> Option<ExtList> {
        self.extension_list.take()
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Pane {
    dock_state: Dock,
    visible: bool,
    width: f64,
    row: u32,
    locked: bool,
    relationship_id: String,
    add_in: AddIn,
    snapshot_resources: Vec<SnapshotResource>,
    extension_list: Option<ExtList>,
}

impl Pane {
    /// Create a right-docked, visible pane with a schema-valid default width.
    #[must_use]
    pub fn new(add_in: AddIn) -> Self {
        Self {
            dock_state: Dock::Right,
            visible: true,
            width: 320.0,
            row: 0,
            locked: false,
            relationship_id: String::new(),
            add_in,
            snapshot_resources: Vec::new(),
            extension_list: None,
        }
    }

    #[must_use]
    pub fn show(mut self, visible: bool) -> Self {
        self.set_visible(visible);
        self
    }

    pub fn set_visible(&mut self, visible: bool) -> &mut Self {
        self.visible = visible;
        self
    }

    pub fn width(mut self, width: f64) -> Result<Self> {
        self.set_width(width)?;
        Ok(self)
    }

    pub fn set_width(&mut self, width: f64) -> Result<&mut Self> {
        if !width.is_finite() || width <= 0.0 {
            return invalid("task-pane width must be finite and positive".into());
        }
        self.width = width;
        Ok(self)
    }

    pub fn dock(mut self, state: impl AsRef<str>) -> Result<Self> {
        self.set_dock(state)?;
        Ok(self)
    }

    pub fn set_dock(&mut self, state: impl AsRef<str>) -> Result<&mut Self> {
        self.dock_state = Dock::parse(state.as_ref())?;
        Ok(self)
    }

    pub fn set_row(&mut self, row: u32) -> &mut Self {
        self.row = row;
        self
    }

    pub fn set_locked(&mut self, locked: bool) -> &mut Self {
        self.locked = locked;
        self
    }

    /// Attach an embedded image using shared storage and a semantic part name.
    pub fn embed(
        mut self,
        part_name: impl AsRef<str>,
        content_type: impl Into<String>,
        data: impl Into<Arc<Vec<u8>>>,
    ) -> Result<Self> {
        self.set_image(part_name, content_type, data)?;
        Ok(self)
    }

    /// Attach or replace the embedded image in place.
    pub fn set_image(
        &mut self,
        part_name: impl AsRef<str>,
        content_type: impl Into<String>,
        data: impl Into<Arc<Vec<u8>>>,
    ) -> Result<&mut Self> {
        let part_name = PackURI::new(part_name.as_ref().to_owned()).map_err(Error::Uri)?;
        if part_name.as_str() == "/" {
            return invalid("snapshot image cannot target the package root".into());
        }
        let content_type = content_type.into();
        validate_image_content_type(&content_type)?;
        let data = data.into();
        if data.len() > MAX_WEB_EXTENSION_SNAPSHOT_BYTES {
            return limit(
                "web extension snapshot bytes",
                MAX_WEB_EXTENSION_SNAPSHOT_BYTES,
                data.len(),
            );
        }
        let linked_id = self
            .add_in
            .snapshot
            .as_ref()
            .and_then(|snapshot| snapshot.linked_relationship_id.as_deref());
        let existing_id = self
            .add_in
            .snapshot
            .as_ref()
            .and_then(|snapshot| snapshot.embedded_relationship_id.as_deref())
            .filter(|id| Some(*id) != linked_id)
            .map(str::to_owned);
        let relationship_id = match existing_id {
            Some(id) => id,
            None => self.next_snapshot_relationship_id("rIdSnapshot")?,
        };
        self.snapshot_resources
            .retain(|resource| resource.relationship_id != relationship_id);
        self.snapshot_resources.push(SnapshotResource {
            relationship_id: relationship_id.clone(),
            target: SnapshotTarget::Internal {
                part_name,
                content_type,
                data,
            },
        });
        self.add_in
            .snapshot
            .get_or_insert_with(Snapshot::default)
            .embedded_relationship_id = Some(relationship_id);
        Ok(self)
    }

    /// Retain an external image link without resolving or contacting it.
    pub fn linked(mut self, target: impl Into<String>) -> Result<Self> {
        self.set_external_link(target)?;
        Ok(self)
    }

    /// Attach or replace an inert external image link in place.
    pub fn set_external_link(&mut self, target: impl Into<String>) -> Result<&mut Self> {
        let target = target.into();
        validate_external_uri_reference(&target)?;
        let embedded_id = self
            .add_in
            .snapshot
            .as_ref()
            .and_then(|snapshot| snapshot.embedded_relationship_id.as_deref());
        let existing_id = self
            .add_in
            .snapshot
            .as_ref()
            .and_then(|snapshot| snapshot.linked_relationship_id.as_deref())
            .filter(|id| Some(*id) != embedded_id)
            .map(str::to_owned);
        let relationship_id = match existing_id {
            Some(id) => id,
            None => self.next_snapshot_relationship_id("rIdSnapshotLink")?,
        };
        self.snapshot_resources
            .retain(|resource| resource.relationship_id != relationship_id);
        self.snapshot_resources.push(SnapshotResource {
            relationship_id: relationship_id.clone(),
            target: SnapshotTarget::External { target },
        });
        self.add_in
            .snapshot
            .get_or_insert_with(Snapshot::default)
            .linked_relationship_id = Some(relationship_id);
        Ok(self)
    }

    /// Attach an internal linked image without exposing its relationship ID.
    pub fn linked_image(
        mut self,
        part_name: impl AsRef<str>,
        content_type: impl Into<String>,
        data: impl Into<Arc<Vec<u8>>>,
    ) -> Result<Self> {
        self.set_linked_image(part_name, content_type, data)?;
        Ok(self)
    }

    /// Attach or replace an internal linked image in place.
    pub fn set_linked_image(
        &mut self,
        part_name: impl AsRef<str>,
        content_type: impl Into<String>,
        data: impl Into<Arc<Vec<u8>>>,
    ) -> Result<&mut Self> {
        let part_name = PackURI::new(part_name.as_ref().to_owned()).map_err(Error::Uri)?;
        if part_name.as_str() == "/" {
            return invalid("linked snapshot image cannot target the package root".into());
        }
        let content_type = content_type.into();
        validate_image_content_type(&content_type)?;
        let data = data.into();
        if data.len() > MAX_WEB_EXTENSION_SNAPSHOT_BYTES {
            return limit(
                "web extension snapshot bytes",
                MAX_WEB_EXTENSION_SNAPSHOT_BYTES,
                data.len(),
            );
        }
        let embedded_id = self
            .add_in
            .snapshot
            .as_ref()
            .and_then(|snapshot| snapshot.embedded_relationship_id.as_deref());
        let existing_id = self
            .add_in
            .snapshot
            .as_ref()
            .and_then(|snapshot| snapshot.linked_relationship_id.as_deref())
            .filter(|id| Some(*id) != embedded_id)
            .map(str::to_owned);
        let relationship_id = match existing_id {
            Some(id) => id,
            None => self.next_snapshot_relationship_id("rIdSnapshotLink")?,
        };
        self.snapshot_resources
            .retain(|resource| resource.relationship_id != relationship_id);
        self.snapshot_resources.push(SnapshotResource {
            relationship_id: relationship_id.clone(),
            target: SnapshotTarget::Internal {
                part_name,
                content_type,
                data,
            },
        });
        self.add_in
            .snapshot
            .get_or_insert_with(Snapshot::default)
            .linked_relationship_id = Some(relationship_id);
        Ok(self)
    }

    /// Set snapshot compression metadata, creating an empty snapshot if needed.
    #[must_use]
    pub fn compress(mut self, compression: Compression) -> Self {
        self.set_compression(Some(compression));
        self
    }

    pub fn set_compression(&mut self, compression: Option<Compression>) -> &mut Self {
        if let Some(compression) = compression {
            self.snapshot_mut().set_compression(Some(compression));
        } else if let Some(snapshot) = self.add_in.snapshot.as_mut() {
            snapshot.set_compression(None);
        }
        self
    }

    /// Append one validated DrawingML effect.
    pub fn effect(mut self, effect: Effect) -> Result<Self> {
        self.push_effect(effect)?;
        Ok(self)
    }

    pub fn push_effect(&mut self, effect: Effect) -> Result<&mut Self> {
        self.snapshot_mut().push_effect(effect)?;
        Ok(self)
    }

    pub fn replace_effect(&mut self, index: usize, effect: Effect) -> Result<Option<Effect>> {
        let Some(snapshot) = self.add_in.snapshot.as_mut() else {
            return Ok(None);
        };
        snapshot.replace_effect(index, effect)
    }

    pub fn remove_effect(&mut self, index: usize) -> Option<Effect> {
        self.add_in
            .snapshot
            .as_mut()
            .and_then(|snapshot| snapshot.remove_effect(index))
    }

    pub fn snapshot_mut(&mut self) -> &mut Snapshot {
        self.add_in.snapshot.get_or_insert_with(Snapshot::default)
    }

    /// Inspect the embedded image without exposing its relationship ID.
    #[must_use]
    pub fn image(&self) -> Option<Image<'_>> {
        let id = self
            .add_in
            .snapshot
            .as_ref()?
            .embedded_relationship_id
            .as_deref()?;
        self.snapshot_resources.iter().find_map(|resource| {
            if resource.relationship_id != id {
                return None;
            }
            match &resource.target {
                SnapshotTarget::Internal {
                    part_name,
                    content_type,
                    data,
                } => Some(Image {
                    part_name,
                    content_type,
                    data,
                }),
                SnapshotTarget::External { .. } => None,
            }
        })
    }

    /// Inspect an internal or inert external link without exposing its relationship ID.
    #[must_use]
    pub fn link(&self) -> Option<Link<'_>> {
        let id = self
            .add_in
            .snapshot
            .as_ref()?
            .linked_relationship_id
            .as_deref()?;
        self.snapshot_resources.iter().find_map(|resource| {
            if resource.relationship_id != id {
                return None;
            }
            match &resource.target {
                SnapshotTarget::External { target } => Some(Link::External(target)),
                SnapshotTarget::Internal {
                    part_name,
                    content_type,
                    data,
                } => Some(Link::Internal(Image {
                    part_name,
                    content_type,
                    data,
                })),
            }
        })
    }

    /// Remove the embedded image, returning whether one existed.
    pub fn clear_image(&mut self) -> bool {
        let Some(id) = self
            .add_in
            .snapshot
            .as_mut()
            .and_then(|snapshot| snapshot.embedded_relationship_id.take())
        else {
            return false;
        };
        let old_len = self.snapshot_resources.len();
        self.snapshot_resources
            .retain(|resource| resource.relationship_id != id);
        old_len != self.snapshot_resources.len()
    }

    /// Remove an internal or external linked image, returning whether one existed.
    pub fn clear_link(&mut self) -> bool {
        let Some(id) = self
            .add_in
            .snapshot
            .as_mut()
            .and_then(|snapshot| snapshot.linked_relationship_id.take())
        else {
            return false;
        };
        let old_len = self.snapshot_resources.len();
        self.snapshot_resources
            .retain(|resource| resource.relationship_id != id);
        old_len != self.snapshot_resources.len()
    }

    pub fn clear_compression(&mut self) -> bool {
        self.add_in
            .snapshot
            .as_mut()
            .and_then(|snapshot| snapshot.compression_state.take())
            .is_some()
    }

    pub fn clear_effects(&mut self) -> bool {
        let Some(snapshot) = self.add_in.snapshot.as_mut() else {
            return false;
        };
        snapshot.clear_effects()
    }

    /// Remove all snapshot XML metadata and embedded/external resources.
    pub fn clear_snapshot(&mut self) -> bool {
        let had_snapshot = self.add_in.snapshot.take().is_some();
        let had_resources = !self.snapshot_resources.is_empty();
        self.snapshot_resources.clear();
        had_snapshot || had_resources
    }

    #[must_use]
    pub const fn visible(&self) -> bool {
        self.visible
    }

    #[must_use]
    pub const fn add_in(&self) -> &AddIn {
        &self.add_in
    }

    pub const fn add_in_mut(&mut self) -> &mut AddIn {
        &mut self.add_in
    }

    #[must_use]
    pub const fn dock_kind(&self) -> &Dock {
        &self.dock_state
    }

    #[must_use]
    pub fn dock_state(&self) -> &str {
        self.dock_state.as_str()
    }

    #[must_use]
    pub const fn pane_width(&self) -> f64 {
        self.width
    }

    #[must_use]
    pub const fn row(&self) -> u32 {
        self.row
    }

    #[must_use]
    pub const fn locked(&self) -> bool {
        self.locked
    }

    #[must_use]
    pub const fn ext(&self) -> Option<&ExtList> {
        self.extension_list.as_ref()
    }

    pub fn set_ext(&mut self, extension: ExtList) -> Result<&mut Self> {
        validate_extension_list(Some(&extension), &[ExtKind::TaskPane])?;
        self.extension_list = Some(extension);
        Ok(self)
    }

    pub fn clear_ext(&mut self) -> Option<ExtList> {
        self.extension_list.take()
    }

    fn next_snapshot_relationship_id(&self, base: &str) -> Result<String> {
        let attempts = self
            .snapshot_resources
            .len()
            .checked_add(2)
            .ok_or(Error::Limit {
                resource: "snapshot relationship IDs",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
        for index in 0..attempts {
            let candidate = if index == 0 {
                base.to_owned()
            } else {
                format!("{base}{index}")
            };
            let used_by_snapshot = self.add_in.snapshot.as_ref().is_some_and(|snapshot| {
                snapshot.embedded_relationship_id.as_deref() == Some(candidate.as_str())
                    || snapshot.linked_relationship_id.as_deref() == Some(candidate.as_str())
            });
            if !used_by_snapshot
                && self
                    .snapshot_resources
                    .iter()
                    .all(|resource| resource.relationship_id != candidate)
            {
                return Ok(candidate);
            }
        }
        Err(Error::Relationship(
            "unable to allocate a snapshot relationship ID".into(),
        ))
    }
}

/// A checked pane selector. Add-in IDs are the semantic primary key; numeric
/// positions are available for ordered document workflows.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Selector<'a> {
    Id(&'a str),
    Index(usize),
}

impl<'a> From<&'a str> for Selector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Id(value)
    }
}

impl From<usize> for Selector<'_> {
    fn from(value: usize) -> Self {
        Self::Index(value)
    }
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct Panes {
    panes: Vec<Pane>,
}

impl Panes {
    #[must_use]
    pub const fn new() -> Self {
        Self { panes: Vec::new() }
    }

    #[must_use]
    pub fn len(&self) -> usize {
        self.panes.len()
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.panes.is_empty()
    }

    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Pane> {
        self.panes.iter()
    }

    /// Look up a pane by semantic add-in ID or checked numeric position.
    #[must_use]
    pub fn get<'a, 'key>(&'a self, selector: impl Into<Selector<'key>>) -> Option<&'a Pane> {
        match selector.into() {
            Selector::Id(id) => self.panes.iter().find(|pane| pane.add_in.id == id),
            Selector::Index(index) => self.panes.get(index),
        }
    }

    /// Edit one pane transactionally while preserving collection-wide invariants.
    ///
    /// Image payloads remain shared through `Arc`; a failed edit leaves the
    /// original pane untouched. Returns `false` when the selector is absent.
    pub fn edit<'key>(
        &mut self,
        selector: impl Into<Selector<'key>>,
        edit: impl FnOnce(&mut Pane) -> Result<()>,
    ) -> Result<bool> {
        let index = match selector.into() {
            Selector::Id(id) => self.panes.iter().position(|pane| pane.add_in.id == id),
            Selector::Index(index) => (index < self.panes.len()).then_some(index),
        };
        let Some(index) = index else {
            return Ok(false);
        };

        let mut candidate = self.panes[index].clone();
        edit(&mut candidate)?;
        if self
            .panes
            .iter()
            .enumerate()
            .any(|(other, pane)| other != index && pane.add_in.id == candidate.add_in.id)
        {
            return invalid(format!("duplicate add-in id '{}'", candidate.add_in.id));
        }
        if self.panes.iter().enumerate().any(|(other, pane)| {
            other != index && pane.relationship_id == candidate.relationship_id
        }) {
            return invalid(format!(
                "duplicate task-pane relationship ID '{}'",
                candidate.relationship_id
            ));
        }
        canonicalize_pane_snapshot_resources(&mut candidate, &self.panes, Some(index))?;
        validate_task_pane(&candidate)?;
        self.panes[index] = candidate;
        Ok(true)
    }

    /// Remove a pane by semantic add-in ID or checked numeric position.
    pub fn remove<'key>(&mut self, selector: impl Into<Selector<'key>>) -> Option<Pane> {
        let index = match selector.into() {
            Selector::Id(id) => self.panes.iter().position(|pane| pane.add_in.id == id)?,
            Selector::Index(index) if index < self.panes.len() => index,
            Selector::Index(_) => return None,
        };
        Some(self.panes.remove(index))
    }

    pub fn push(&mut self, mut pane: Pane) -> Result<&mut Self> {
        if self.panes.len() >= MAX_WEB_EXTENSION_ITEMS {
            return limit(
                "web extension panes",
                MAX_WEB_EXTENSION_ITEMS,
                self.panes.len().saturating_add(1),
            );
        }
        if self
            .panes
            .iter()
            .any(|value| value.add_in.id == pane.add_in.id)
        {
            return invalid(format!("duplicate add-in id '{}'", pane.add_in.id));
        }
        canonicalize_pane_snapshot_resources(&mut pane, &self.panes, None)?;
        if pane.relationship_id.is_empty()
            || self
                .panes
                .iter()
                .any(|value| value.relationship_id == pane.relationship_id)
        {
            pane.relationship_id = self.next_relationship_id()?;
        }
        validate_task_pane(&pane)?;
        self.panes.push(pane);
        Ok(self)
    }

    fn next_relationship_id(&self) -> Result<String> {
        let attempts = self.panes.len().checked_add(1).ok_or(Error::Limit {
            resource: "web extension pane relationship IDs",
            max: usize::MAX,
            actual: usize::MAX,
        })?;
        for index in 1..=attempts {
            let candidate = format!("rIdAddIn{index}");
            if self
                .panes
                .iter()
                .all(|pane| pane.relationship_id != candidate)
            {
                return Ok(candidate);
            }
        }
        Err(Error::Relationship(
            "unable to allocate an add-in relationship ID".into(),
        ))
    }
}

fn canonicalize_pane_snapshot_resources(
    pane: &mut Pane,
    existing_panes: &[Pane],
    skip_index: Option<usize>,
) -> Result<()> {
    for index in 0..pane.snapshot_resources.len() {
        let (previous, incoming) = pane.snapshot_resources.split_at_mut(index);
        let incoming = &mut incoming[0];
        for existing in previous {
            reconcile_snapshot_resource(incoming, existing, "another resource in the same pane")?;
        }
    }

    for incoming in &mut pane.snapshot_resources {
        for (pane_index, existing_pane) in existing_panes.iter().enumerate() {
            if skip_index == Some(pane_index) {
                continue;
            }
            for existing in &existing_pane.snapshot_resources {
                reconcile_snapshot_resource(incoming, existing, "an existing pane resource")?;
            }
        }
    }
    Ok(())
}

fn reconcile_snapshot_resource(
    incoming: &mut SnapshotResource,
    existing: &SnapshotResource,
    conflict_scope: &str,
) -> Result<()> {
    let (
        SnapshotTarget::Internal {
            part_name,
            content_type,
            data,
        },
        SnapshotTarget::Internal {
            part_name: existing_name,
            content_type: existing_content_type,
            data: existing_data,
        },
    ) = (&mut incoming.target, &existing.target)
    else {
        return Ok(());
    };
    if !part_names_conflict(part_name, existing_name) {
        return Ok(());
    }
    if part_name.is_equivalent_to(existing_name)
        && content_type == existing_content_type
        && data.as_slice() == existing_data.as_slice()
    {
        *part_name = existing_name.clone();
        return Ok(());
    }
    invalid(format!(
        "snapshot part '{}' conflicts with {conflict_scope}",
        part_name.as_str()
    ))
}

#[derive(Debug, Default)]
struct OperationBudget {
    xml_bytes: usize,
    string_bytes: usize,
}

impl OperationBudget {
    fn charge_xml(&mut self, bytes: usize, limits: &Limits) -> Result<()> {
        self.xml_bytes = self.xml_bytes.checked_add(bytes).ok_or(Error::Limit {
            resource: "aggregate web extension XML bytes",
            max: limits.total_xml_bytes,
            actual: usize::MAX,
        })?;
        if self.xml_bytes > limits.total_xml_bytes {
            return limit(
                "aggregate web extension XML bytes",
                limits.total_xml_bytes,
                self.xml_bytes,
            );
        }
        Ok(())
    }

    fn charge_strings(&mut self, bytes: usize, limits: &Limits) -> Result<()> {
        self.string_bytes = self.string_bytes.checked_add(bytes).ok_or(Error::Limit {
            resource: "retained web extension string bytes",
            max: limits.total_string_bytes,
            actual: usize::MAX,
        })?;
        if self.string_bytes > limits.total_string_bytes {
            return limit(
                "retained web extension string bytes",
                limits.total_string_bytes,
                self.string_bytes,
            );
        }
        Ok(())
    }

    fn charge_authored(&mut self, xml: &[u8], limits: &Limits) -> Result<()> {
        self.charge_xml(xml.len(), limits)?;
        self.charge_strings(xml.len(), limits)
    }

    fn charge_metadata(&mut self, bytes: usize, copies: usize, limits: &Limits) -> Result<()> {
        let retained = bytes.checked_mul(copies).ok_or(Error::Limit {
            resource: "indexed web extension package metadata bytes",
            max: limits.total_string_bytes,
            actual: usize::MAX,
        })?;
        self.charge_strings(retained, limits)
    }
}

/// Parse one MS-OWEXML web extension part after bounded MCE preprocessing.
#[cfg(test)]
fn parse_add_in(xml: &[u8]) -> Result<AddIn> {
    parse_add_in_with(xml, &Limits::standard())
}

fn parse_add_in_with(xml: &[u8], limits: &Limits) -> Result<AddIn> {
    let mut budget = OperationBudget::default();
    parse_add_in_with_budget(xml, limits, &mut budget)
}

fn parse_add_in_with_budget(
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
fn parse_panes(xml: &[u8]) -> Result<Vec<ParsedPane>> {
    parse_panes_with(xml, &Limits::standard())
}

fn parse_panes_with(xml: &[u8], limits: &Limits) -> Result<Vec<ParsedPane>> {
    let mut budget = OperationBudget::default();
    parse_panes_with_budget(xml, limits, &mut budget)
}

fn parse_panes_with_budget(
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
struct ParsedPane {
    dock_state: Dock,
    visible: bool,
    width: f64,
    row: u32,
    locked: bool,
    relationship_id: String,
    extension_list: Option<ExtList>,
}

fn has_task_panes_relationship(package: &OpcPackage, limits: &Limits) -> Result<bool> {
    let relationships = package.rels().len();
    if relationships > limits.package_relationships {
        return limit(
            "package relationships",
            limits.package_relationships,
            relationships,
        );
    }
    Ok(package
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP))
}

/// Resolve and validate the complete package graph with safe default limits.
pub fn load(package: &OpcPackage) -> Result<Option<Panes>> {
    load_with(package, &Limits::standard())
}

/// Resolve and validate the complete package graph with explicit limits.
pub fn load_with(package: &OpcPackage, limits: &Limits) -> Result<Option<Panes>> {
    if !has_task_panes_relationship(package, limits)? {
        return Ok(None);
    }
    let mut budget = OperationBudget::default();
    let index = PackageGraphIndex::build(package, limits, &mut budget)?;
    load_with_index_budget(package, limits, &index, &mut budget)
}

fn load_with_index_budget(
    package: &OpcPackage,
    limits: &Limits,
    index: &PackageGraphIndex,
    budget: &mut OperationBudget,
) -> Result<Option<Panes>> {
    let relationships: Vec<_> = package
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
        .collect();
    if relationships.is_empty() {
        return Ok(None);
    }
    if relationships.len() != 1 {
        return Err(Error::Relationship(
            "package has multiple web extension task-pane relationships".into(),
        ));
    }
    let relationship = relationships[0];
    if relationship.is_external() {
        return Err(Error::Relationship(
            "task-pane relationship must be internal".into(),
        ));
    }
    let target = checked_internal_target(relationship, "task-pane")?;
    let part_name = index
        .canonical(&target)
        .ok_or_else(|| Error::Missing(format!("task-pane part '{}'", target.as_str())))?;
    let task_panes_part = package.get_part(part_name).map_err(|error| {
        Error::Missing(format!("task-pane part '{}': {error}", part_name.as_str()))
    })?;
    require_content_type(task_panes_part, TASK_PANES_CONTENT_TYPE)?;
    let parsed_panes = parse_panes_with_budget(task_panes_part.blob(), limits, budget)?;

    let referenced_ids: HashSet<&str> = parsed_panes
        .iter()
        .map(|pane| pane.relationship_id.as_str())
        .collect();
    if referenced_ids.len() != parsed_panes.len() {
        return invalid("task panes contain duplicate relationship IDs".into());
    }
    for child_relationship in task_panes_part.rels().iter() {
        if child_relationship.reltype() != ADD_IN_RELATIONSHIP {
            return Err(Error::Relationship(format!(
                "task-pane part has forbidden relationship '{}' of type '{}'",
                child_relationship.r_id(),
                child_relationship.reltype()
            )));
        }
        if !referenced_ids.contains(child_relationship.r_id()) {
            return Err(Error::Relationship(format!(
                "task-pane part has unreferenced relationship '{}'",
                child_relationship.r_id()
            )));
        }
    }

    let mut panes = Vec::with_capacity(parsed_panes.len());
    let mut total_snapshot_bytes = 0usize;
    let mut snapshot_names = HashSet::new();
    let mut extension_names = HashSet::new();
    for pane in parsed_panes {
        let child_relationship = task_panes_part
            .rels()
            .get(&pane.relationship_id)
            .ok_or_else(|| {
                Error::Relationship(format!(
                    "task pane references missing relationship '{}'",
                    pane.relationship_id
                ))
            })?;
        if child_relationship.is_external() {
            return Err(Error::Relationship(format!(
                "web extension relationship '{}' must be internal",
                pane.relationship_id
            )));
        }
        let extension_target = checked_internal_target(child_relationship, "add-in")?;
        let extension_name = index.canonical(&extension_target).ok_or_else(|| {
            Error::Missing(format!(
                "web extension part '{}'",
                extension_target.as_str()
            ))
        })?;
        let extension_part = package.get_part(extension_name).map_err(|error| {
            Error::Missing(format!(
                "web extension part '{}': {error}",
                extension_name.as_str()
            ))
        })?;
        let extension_name = extension_part.partname().clone();
        if !extension_names.insert(fold_part_name(&extension_name)) {
            return Err(Error::Relationship(format!(
                "multiple task panes target web extension part '{}'",
                extension_name.as_str()
            )));
        }
        require_content_type(extension_part, ADD_IN_CONTENT_TYPE)?;
        let add_in = parse_add_in_with_budget(extension_part.blob(), limits, budget)?;
        let snapshot_resources = load_snapshot_resources(
            package,
            extension_part,
            &add_in,
            &mut total_snapshot_bytes,
            &mut snapshot_names,
            limits,
            index,
        )?;
        panes.push(Pane {
            dock_state: pane.dock_state,
            visible: pane.visible,
            width: pane.width,
            row: pane.row,
            locked: pane.locked,
            relationship_id: pane.relationship_id,
            add_in,
            snapshot_resources,
            extension_list: pane.extension_list,
        });
    }
    let panes = Panes { panes };
    validate_panes(&panes, limits)?;
    Ok(Some(panes))
}

/// Create or replace the package-level persisted task-pane graph.
///
/// Add-in references, bindings, properties, and snapshot resources are stored
/// as inert data. External snapshot links are never contacted.
pub fn put(package: &mut OpcPackage, panes: Panes, conformance: Conformance) -> Result<()> {
    put_with(package, panes, conformance, &Limits::standard())
}

/// Create or replace the task-pane graph with explicit resource limits.
pub fn put_with(
    package: &mut OpcPackage,
    task_panes: Panes,
    conformance: Conformance,
    limits: &Limits,
) -> Result<()> {
    let mut budget = OperationBudget::default();
    let index = PackageGraphIndex::build(package, limits, &mut budget)?;
    charge_authored_metadata(&task_panes, &mut budget, limits)?;
    validate_panes(&task_panes, limits)?;
    let task_panes_xml = write_panes_with(&task_panes, conformance, limits)?;
    budget.charge_authored(&task_panes_xml, limits)?;
    let existing = existing_web_extension_graph(package, limits, &index, &mut budget)?;
    let mut allocation_probes = 0usize;
    let task_panes_name = match existing.as_ref() {
        Some(graph) => graph.task_panes_name.clone(),
        None => next_task_panes_part_name(&index, limits, &mut allocation_probes)?,
    };
    let mut reserved_names = BTreeSet::new();
    reserved_names.insert(fold_part_name(&task_panes_name));
    let mut planned = Vec::with_capacity(task_panes.panes.len() + 1);
    let mut planned_by_name = HashMap::with_capacity(task_panes.panes.len() + 1);
    let mut task_relationships = Vec::with_capacity(task_panes.panes.len());
    let mut total_snapshot_bytes = 0usize;
    let mut counted_snapshot_parts = HashSet::new();
    let existing_extensions = existing
        .as_ref()
        .map(|graph| &graph.extensions_by_relationship);

    for (pane_index, pane) in task_panes.panes.iter().enumerate() {
        let extension_name = match existing_extensions
            .and_then(|extensions| extensions.get(&pane.relationship_id))
        {
            Some(name) => name.clone(),
            None => next_web_extension_part_name(
                &index,
                &reserved_names,
                pane_index + 1,
                limits,
                &mut allocation_probes,
            )?,
        };
        let extension_key = fold_part_name(&extension_name);
        if folded_name_conflicts(&reserved_names, &extension_key) {
            return invalid(format!(
                "multiple task panes target web extension part '{}'",
                extension_name.as_str()
            ));
        }
        reserved_names.insert(extension_key);
        let extension_xml = write_add_in_with(&pane.add_in, conformance, limits)?;
        budget.charge_authored(&extension_xml, limits)?;
        let mut relationships = Vec::with_capacity(pane.snapshot_resources.len());
        for resource in &pane.snapshot_resources {
            let (target, external) = match &resource.target {
                SnapshotTarget::Internal {
                    part_name,
                    content_type,
                    data,
                } => {
                    let part_key = fold_part_name(part_name);
                    let already_counted = counted_snapshot_parts.contains(&part_key);
                    if folded_name_conflicts(&reserved_names, &part_key) && !already_counted {
                        return invalid(format!(
                            "snapshot part '{}' conflicts with another authored part",
                            part_name.as_str()
                        ));
                    }
                    reserved_names.insert(part_key.clone());
                    if counted_snapshot_parts.insert(part_key) {
                        total_snapshot_bytes = total_snapshot_bytes
                            .checked_add(data.len())
                            .ok_or_else(|| {
                                Error::Invalid("aggregate snapshot byte count overflow".into())
                            })?;
                        if total_snapshot_bytes > limits.total_image_bytes {
                            return limit(
                                "aggregate web extension snapshot bytes",
                                limits.total_image_bytes,
                                total_snapshot_bytes,
                            );
                        }
                    }
                    add_or_match_planned_part(
                        &mut planned,
                        &mut planned_by_name,
                        PlannedPart {
                            name: part_name.clone(),
                            content_type: content_type.clone(),
                            data: data.clone(),
                            relationships: Vec::new(),
                        },
                    )?;
                    (part_name.relative_ref(extension_name.base_uri()), false)
                },
                SnapshotTarget::External { target } => (target.clone(), true),
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
            &mut planned_by_name,
            PlannedPart {
                name: extension_name.clone(),
                content_type: ADD_IN_CONTENT_TYPE.into(),
                data: Arc::new(extension_xml),
                relationships,
            },
        )?;
        task_relationships.push(PlannedRelationship {
            id: pane.relationship_id.clone(),
            relationship_type: ADD_IN_RELATIONSHIP.into(),
            target: extension_name.relative_ref(task_panes_name.base_uri()),
            external: false,
        });
    }
    add_or_match_planned_part(
        &mut planned,
        &mut planned_by_name,
        PlannedPart {
            name: task_panes_name.clone(),
            content_type: TASK_PANES_CONTENT_TYPE.into(),
            data: Arc::new(task_panes_xml),
            relationships: task_relationships,
        },
    )?;

    let old_parts = existing
        .as_ref()
        .map_or(&[][..], |graph| graph.owned_parts.as_slice());
    let protected = existing.as_ref().map_or_else(HashSet::new, |graph| {
        index.protected_closure(&graph.owned_parts, &graph.root_relationship_id)
    });
    preflight_planned_parts(package, &index, &planned, old_parts, &protected)?;
    if existing
        .as_ref()
        .is_some_and(|graph| graph_matches_plan(package, &planned, graph, &reserved_names))
    {
        return Ok(());
    }
    let deletions = planned_deletions(old_parts, &reserved_names, &protected, limits)?;
    validate_plan_counts(package, &index, &planned, existing.as_ref(), limits)?;
    let root_relationship_id = existing
        .as_ref()
        .map(|graph| graph.root_relationship_id.clone())
        .map_or_else(|| next_package_relationship_id(package, limits), Ok)?;
    install_planned_parts(package, planned)?;

    if existing.is_some() {
        package.rels_mut().remove(&root_relationship_id);
    }
    package.rels_mut().add_relationship(
        TASK_PANES_RELATIONSHIP.into(),
        task_panes_name.as_str().trim_start_matches('/').into(),
        root_relationship_id,
        false,
    );

    for name in deletions {
        package.remove_part(&name);
    }
    package.unsign();
    Ok(())
}

/// Remove the package-level task-pane relationship and graph.
///
/// Parts still referenced elsewhere remain in the package.
pub fn remove(package: &mut OpcPackage) -> Result<bool> {
    remove_with(package, &Limits::standard())
}

/// Remove the task-pane graph with explicit package graph and deletion ceilings.
pub fn remove_with(package: &mut OpcPackage, limits: &Limits) -> Result<bool> {
    if !has_task_panes_relationship(package, limits)? {
        return Ok(false);
    }
    let mut budget = OperationBudget::default();
    let index = PackageGraphIndex::build(package, limits, &mut budget)?;
    let Some(existing) = existing_web_extension_graph(package, limits, &index, &mut budget)? else {
        return Ok(false);
    };
    let protected = index.protected_closure(&existing.owned_parts, &existing.root_relationship_id);
    let deletions = planned_deletions(&existing.owned_parts, &BTreeSet::new(), &protected, limits)?;
    package.rels_mut().remove(&existing.root_relationship_id);
    for name in deletions {
        package.remove_part(&name);
    }
    package.unsign();
    Ok(true)
}

#[derive(Debug)]
struct ExistingAddInGraph {
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
    data: Arc<Vec<u8>>,
    relationships: Vec<PlannedRelationship>,
}

#[derive(Debug)]
struct IndexedInbound {
    source: Option<usize>,
    relationship_id: String,
}

#[derive(Debug)]
struct IndexedPart {
    name: PackURI,
    outbound: Vec<usize>,
    inbound: Vec<IndexedInbound>,
}

/// One bounded, ASCII-case-folded view of package membership and internal edges.
#[derive(Debug)]
struct PackageGraphIndex {
    parts: Vec<IndexedPart>,
    by_folded: HashMap<String, usize>,
    occupied: BTreeSet<String>,
    relationships: usize,
}

impl PackageGraphIndex {
    fn build(package: &OpcPackage, limits: &Limits, budget: &mut OperationBudget) -> Result<Self> {
        let part_count = package.part_count();
        if part_count > limits.package_parts {
            return limit("package parts", limits.package_parts, part_count);
        }
        let mut parts: Vec<IndexedPart> = Vec::with_capacity(part_count);
        let mut by_folded = HashMap::with_capacity(part_count);
        let mut occupied = BTreeSet::new();
        for part in package.iter_parts() {
            let metadata_bytes = part
                .partname()
                .as_str()
                .len()
                .checked_add(part.content_type().len())
                .ok_or(Error::Limit {
                    resource: "indexed web extension package metadata bytes",
                    max: limits.total_string_bytes,
                    actual: usize::MAX,
                })?;
            budget.charge_metadata(metadata_bytes, 4, limits)?;
            let folded = fold_part_name(part.partname());
            if let Some(index) = by_folded.insert(folded.clone(), parts.len()) {
                return invalid(format!(
                    "ASCII-case-equivalent package parts '{}' and '{}' coexist",
                    parts[index].name.as_str(),
                    part.partname().as_str()
                ));
            }
            occupied.insert(folded);
            parts.push(IndexedPart {
                name: part.partname().clone(),
                outbound: Vec::new(),
                inbound: Vec::new(),
            });
        }

        let mut value = Self {
            parts,
            by_folded,
            occupied,
            relationships: 0,
        };
        for relationship in package.rels().iter() {
            value.record_relationship(None, relationship, limits, budget)?;
        }
        for part in package.iter_parts() {
            let source = value
                .index_of(part.partname())
                .ok_or_else(|| Error::Missing(part.partname().to_string()))?;
            for relationship in part.rels().iter() {
                value.record_relationship(Some(source), relationship, limits, budget)?;
            }
        }
        Ok(value)
    }

    fn record_relationship(
        &mut self,
        source: Option<usize>,
        relationship: &litchi_opc::Relationship,
        limits: &Limits,
        budget: &mut OperationBudget,
    ) -> Result<()> {
        let metadata_bytes = relationship
            .r_id()
            .len()
            .checked_add(relationship.reltype().len())
            .and_then(|bytes| bytes.checked_add(relationship.target_ref().len()))
            .ok_or(Error::Limit {
                resource: "indexed web extension package metadata bytes",
                max: limits.total_string_bytes,
                actual: usize::MAX,
            })?;
        budget.charge_metadata(metadata_bytes, 3, limits)?;
        self.relationships = self.relationships.checked_add(1).ok_or(Error::Limit {
            resource: "package relationships",
            max: limits.package_relationships,
            actual: usize::MAX,
        })?;
        if self.relationships > limits.package_relationships {
            return limit(
                "package relationships",
                limits.package_relationships,
                self.relationships,
            );
        }
        if relationship.is_external() {
            return Ok(());
        }
        let Ok(target) = relationship.target_partname() else {
            // Web graph relationships are rejected with context by their callers.
            return Ok(());
        };
        let Some(target) = self.index_of(&target) else {
            return Ok(());
        };
        if let Some(source) = source {
            self.parts[source].outbound.push(target);
        }
        self.parts[target].inbound.push(IndexedInbound {
            source,
            relationship_id: relationship.r_id().to_owned(),
        });
        Ok(())
    }

    fn index_of(&self, name: &PackURI) -> Option<usize> {
        self.by_folded.get(&fold_part_name(name)).copied()
    }

    fn canonical(&self, name: &PackURI) -> Option<&PackURI> {
        self.index_of(name).map(|index| &self.parts[index].name)
    }

    fn contains(&self, name: &PackURI) -> bool {
        self.index_of(name).is_some()
    }

    fn conflicts(&self, candidate: &PackURI) -> bool {
        let folded = fold_part_name(candidate);
        if self.occupied.contains(&folded) {
            return true;
        }
        let mut ancestor = folded.as_str();
        while let Some(index) = ancestor.rfind('/') {
            if index == 0 {
                break;
            }
            ancestor = &ancestor[..index];
            if self.occupied.contains(ancestor) {
                return true;
            }
        }
        let descendant_prefix = format!("{folded}/");
        self.occupied
            .range(descendant_prefix.clone()..)
            .next()
            .is_some_and(|name| name.starts_with(&descendant_prefix))
    }

    fn protected_closure(
        &self,
        owned_parts: &[PackURI],
        root_relationship_id: &str,
    ) -> HashSet<String> {
        let owned: HashSet<_> = owned_parts
            .iter()
            .filter_map(|name| self.index_of(name))
            .collect();
        let mut queue = VecDeque::new();
        let mut protected = HashSet::new();
        for &index in &owned {
            let has_external_ingress =
                self.parts[index]
                    .inbound
                    .iter()
                    .any(|inbound| match inbound.source {
                        None => inbound.relationship_id != root_relationship_id,
                        Some(source) => !owned.contains(&source),
                    });
            if has_external_ingress && protected.insert(index) {
                queue.push_back(index);
            }
        }
        while let Some(source) = queue.pop_front() {
            for &target in &self.parts[source].outbound {
                if owned.contains(&target) && protected.insert(target) {
                    queue.push_back(target);
                }
            }
        }
        protected
            .into_iter()
            .map(|index| fold_part_name(&self.parts[index].name))
            .collect()
    }
}

fn fold_part_name(name: &PackURI) -> String {
    name.as_str().to_ascii_lowercase()
}

fn existing_web_extension_graph(
    package: &OpcPackage,
    limits: &Limits,
    index: &PackageGraphIndex,
    budget: &mut OperationBudget,
) -> Result<Option<ExistingAddInGraph>> {
    let Some(loaded) = load_with_index_budget(package, limits, index, budget)? else {
        return Ok(None);
    };
    let relationship = package
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
        .ok_or_else(|| {
            Error::Relationship("loaded task panes have no package relationship".into())
        })?;
    let task_panes_target = checked_internal_target(relationship, "task-pane")?;
    let task_panes_name = index
        .canonical(&task_panes_target)
        .ok_or_else(|| Error::Missing(task_panes_target.to_string()))?
        .clone();
    let task_panes_part = package.get_part(&task_panes_name)?;
    let mut extensions_by_relationship = HashMap::with_capacity(loaded.panes.len());
    let mut owned = HashSet::new();
    owned.insert(task_panes_name.clone());
    for pane in loaded.panes {
        let child_relationship = task_panes_part
            .rels()
            .get(&pane.relationship_id)
            .ok_or_else(|| {
                Error::Relationship(format!(
                    "task pane references missing relationship '{}'",
                    pane.relationship_id
                ))
            })?;
        let extension_target = checked_internal_target(child_relationship, "add-in")?;
        let extension_name = index
            .canonical(&extension_target)
            .ok_or_else(|| Error::Missing(extension_target.to_string()))?
            .clone();
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
            if let SnapshotTarget::Internal { part_name, .. } = resource.target {
                owned.insert(part_name);
            }
        }
    }
    let mut owned_parts: Vec<_> = owned.into_iter().collect();
    owned_parts.sort_by(|left, right| left.as_str().cmp(right.as_str()));
    Ok(Some(ExistingAddInGraph {
        root_relationship_id: relationship.r_id().to_owned(),
        task_panes_name,
        extensions_by_relationship,
        owned_parts,
    }))
}

fn next_web_extension_part_name(
    index: &PackageGraphIndex,
    reserved: &BTreeSet<String>,
    preferred_index: usize,
    limits: &Limits,
    allocation_probes: &mut usize,
) -> Result<PackURI> {
    let mut offset = 0usize;
    loop {
        charge_allocation_probe(allocation_probes, limits)?;
        let part_number = preferred_index
            .checked_add(offset)
            .ok_or_else(|| Error::Invalid("web extension part index overflow".into()))?;
        let candidate = PackURI::new(format!("/webextensions/webextension{part_number}.xml"))
            .map_err(Error::Uri)?;
        if !folded_name_conflicts(reserved, &fold_part_name(&candidate))
            && !index.conflicts(&candidate)
        {
            return Ok(candidate);
        }
        offset = offset
            .checked_add(1)
            .ok_or_else(|| Error::Invalid("web extension part index overflow".into()))?;
    }
}

fn next_task_panes_part_name(
    index: &PackageGraphIndex,
    limits: &Limits,
    allocation_probes: &mut usize,
) -> Result<PackURI> {
    let mut attempt = 1usize;
    loop {
        charge_allocation_probe(allocation_probes, limits)?;
        let suffix = if attempt == 1 {
            String::new()
        } else {
            attempt.to_string()
        };
        let candidate =
            PackURI::new(format!("/webextensions/taskpanes{suffix}.xml")).map_err(Error::Uri)?;
        if !index.conflicts(&candidate) {
            return Ok(candidate);
        }
        attempt = attempt
            .checked_add(1)
            .ok_or_else(|| Error::Invalid("task-pane part index overflow".into()))?;
    }
}

fn charge_allocation_probe(probes: &mut usize, limits: &Limits) -> Result<()> {
    let actual = probes.saturating_add(1);
    if *probes >= limits.part_allocations {
        return limit(
            "web extension part-name allocation probes",
            limits.part_allocations,
            actual,
        );
    }
    *probes = actual;
    Ok(())
}

fn next_package_relationship_id(package: &OpcPackage, limits: &Limits) -> Result<String> {
    let attempts = package.rels().len().checked_add(1).ok_or(Error::Limit {
        resource: "web extension relationship IDs",
        max: limits.part_allocations,
        actual: usize::MAX,
    })?;
    let attempts = attempts.min(limits.part_allocations);
    for index in 1..=attempts {
        let candidate = format!("rIdPanes{index}");
        if package.rels().get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(Error::Relationship(
        "unable to allocate a task-pane relationship ID".into(),
    ))
}

fn add_or_match_planned_part(
    planned: &mut Vec<PlannedPart>,
    by_name: &mut HashMap<String, usize>,
    part: PlannedPart,
) -> Result<()> {
    let key = fold_part_name(&part.name);
    if let Some(existing) = by_name.get(&key).and_then(|index| planned.get(*index)) {
        if existing == &part {
            return Ok(());
        }
        return invalid(format!(
            "conflicting authored resources target '{}'",
            part.name.as_str()
        ));
    }
    by_name.insert(key, planned.len());
    planned.push(part);
    Ok(())
}

fn preflight_planned_parts(
    package: &OpcPackage,
    index: &PackageGraphIndex,
    planned: &[PlannedPart],
    old_parts: &[PackURI],
    protected: &HashSet<String>,
) -> Result<()> {
    let old_names: HashSet<_> = old_parts.iter().map(fold_part_name).collect();
    let mut planned_names = BTreeSet::new();
    for part in planned {
        let folded = fold_part_name(&part.name);
        if folded_name_conflicts(&planned_names, &folded) {
            return invalid(format!(
                "authored part '{}' conflicts with another planned part",
                part.name.as_str()
            ));
        }
        planned_names.insert(folded.clone());
        if let Some(canonical) = index.canonical(&part.name) {
            if canonical != &part.name {
                return invalid(format!(
                    "authored part name '{}' differs in case from canonical package part '{}'",
                    part.name.as_str(),
                    canonical.as_str()
                ));
            }
            if !old_names.contains(&folded) {
                return invalid(format!(
                    "authored web extension part '{}' already exists outside the replaced graph",
                    part.name.as_str()
                ));
            }
            let existing = package.get_part(canonical)?;
            if existing.content_type() != part.content_type {
                return invalid(format!(
                    "cannot change content type of existing part '{}'",
                    part.name.as_str()
                ));
            }
            if !planned_part_matches_existing(existing, part) && protected.contains(&folded) {
                return Err(Error::Relationship(format!(
                    "cannot replace protected shared web extension part '{}'",
                    part.name.as_str()
                )));
            }
        } else if index.conflicts(&part.name) {
            return invalid(format!(
                "authored part '{}' conflicts with an existing package part",
                part.name.as_str()
            ));
        }
    }
    Ok(())
}

fn folded_name_conflicts(names: &BTreeSet<String>, candidate: &str) -> bool {
    if names.contains(candidate) {
        return true;
    }
    let mut ancestor = candidate;
    while let Some(index) = ancestor.rfind('/') {
        if index == 0 {
            break;
        }
        ancestor = &ancestor[..index];
        if names.contains(ancestor) {
            return true;
        }
    }
    let prefix = format!("{candidate}/");
    names
        .range(prefix.clone()..)
        .next()
        .is_some_and(|name| name.starts_with(&prefix))
}

fn part_names_conflict(left: &PackURI, right: &PackURI) -> bool {
    let left = fold_part_name(left);
    let right = fold_part_name(right);
    left == right
        || left
            .strip_prefix(&right)
            .is_some_and(|suffix| suffix.starts_with('/'))
        || right
            .strip_prefix(&left)
            .is_some_and(|suffix| suffix.starts_with('/'))
}

fn planned_deletions(
    old_parts: &[PackURI],
    retained: &BTreeSet<String>,
    protected: &HashSet<String>,
    limits: &Limits,
) -> Result<Vec<PackURI>> {
    let mut deletions = Vec::new();
    for name in old_parts {
        let folded = fold_part_name(name);
        if protected.contains(&folded) || retained.contains(&folded) {
            continue;
        }
        deletions.push(name.clone());
    }
    if deletions.len() > limits.part_deletions {
        return limit(
            "web extension part deletions",
            limits.part_deletions,
            deletions.len(),
        );
    }
    Ok(deletions)
}

fn validate_plan_counts(
    package: &OpcPackage,
    index: &PackageGraphIndex,
    planned: &[PlannedPart],
    existing_graph: Option<&ExistingAddInGraph>,
    limits: &Limits,
) -> Result<()> {
    let new_parts = planned
        .iter()
        .filter(|part| !index.contains(&part.name))
        .count();
    if new_parts > limits.part_allocations {
        return limit(
            "web extension part allocations",
            limits.part_allocations,
            new_parts,
        );
    }
    let peak_parts = index
        .parts
        .len()
        .checked_add(new_parts)
        .ok_or(Error::Limit {
            resource: "package parts",
            max: limits.package_parts,
            actual: usize::MAX,
        })?;
    if peak_parts > limits.package_parts {
        return limit("package parts", limits.package_parts, peak_parts);
    }

    let mut relationships = index.relationships;
    for part in planned {
        if let Some(canonical) = index.canonical(&part.name) {
            relationships = relationships
                .checked_sub(package.get_part(canonical)?.rels().len())
                .ok_or_else(|| Error::Invalid("relationship count underflow".into()))?;
        }
        relationships =
            relationships
                .checked_add(part.relationships.len())
                .ok_or(Error::Limit {
                    resource: "package relationships",
                    max: limits.package_relationships,
                    actual: usize::MAX,
                })?;
    }
    if existing_graph.is_none() {
        relationships = relationships.checked_add(1).ok_or(Error::Limit {
            resource: "package relationships",
            max: limits.package_relationships,
            actual: usize::MAX,
        })?;
    }
    if relationships > limits.package_relationships {
        return limit(
            "package relationships",
            limits.package_relationships,
            relationships,
        );
    }
    Ok(())
}

fn graph_matches_plan(
    package: &OpcPackage,
    planned: &[PlannedPart],
    existing: &ExistingAddInGraph,
    retained: &BTreeSet<String>,
) -> bool {
    existing.owned_parts.len() == retained.len()
        && existing
            .owned_parts
            .iter()
            .all(|name| retained.contains(&fold_part_name(name)))
        && planned.iter().all(|part| {
            package
                .get_part(&part.name)
                .is_ok_and(|existing_part| planned_part_matches_existing(existing_part, part))
        })
}

fn planned_part_matches_existing(existing: &dyn Part, planned: &PlannedPart) -> bool {
    existing.content_type() == planned.content_type
        && existing.blob() == planned.data.as_slice()
        && existing.rels().len() == planned.relationships.len()
        && planned.relationships.iter().all(|planned_relationship| {
            existing
                .rels()
                .get(&planned_relationship.id)
                .is_some_and(|relationship| {
                    relationship.reltype() == planned_relationship.relationship_type
                        && relationship.target_ref() == planned_relationship.target
                        && relationship.is_external() == planned_relationship.external
                })
        })
}

fn install_planned_parts(package: &mut OpcPackage, planned: Vec<PlannedPart>) -> Result<()> {
    let staged = planned
        .into_iter()
        .map(planned_blob_part)
        .collect::<Result<Vec<_>>>()?;
    // Preflight and relationship construction complete before this point. The
    // remaining map replacements are infallible, so callers never observe a
    // partially installed Web Extensions graph on error.
    for part in staged {
        package.add_part(Box::new(part));
    }
    Ok(())
}

fn planned_blob_part(part: PlannedPart) -> Result<BlobPart> {
    let mut value = BlobPart::new_shared(part.name, part.content_type, part.data);
    for relationship in part.relationships {
        value.rels_mut().try_add_relationship(
            relationship.relationship_type,
            relationship.target,
            relationship.id,
            if relationship.external {
                TargetMode::External
            } else {
                TargetMode::Internal
            },
        )?;
    }
    Ok(value)
}

/// Deterministically serialize a single web extension part.
#[cfg(test)]
fn write_add_in(extension: &AddIn, conformance: Conformance) -> Result<Vec<u8>> {
    write_add_in_with(extension, conformance, &Limits::standard())
}

fn write_add_in_with(
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
fn write_panes(task_panes: &Panes, conformance: Conformance) -> Result<Vec<u8>> {
    write_panes_with(task_panes, conformance, &Limits::standard())
}

fn write_panes_with(
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

fn parse_task_pane(node: &Node, document: &XmlDocument) -> Result<ParsedPane> {
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

fn parse_store_reference(node: &Node, document: &XmlDocument) -> Result<Reference> {
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
    Ok(Reference {
        id: required_attr(node, "", "id")?.to_owned(),
        version: required_attr(node, "", "version")?.to_owned(),
        catalog: attr(node, "", "store").map(str::to_owned),
        store: attr(node, "", "storeType")
            .map(Store::parse)
            .transpose()?
            .unwrap_or_default(),
        extension_list: children
            .first()
            .map(|node| ExtList::from_node(node, document))
            .transpose()?,
    })
}

fn parse_property(node: &Node) -> Result<Property> {
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

fn parse_binding(node: &Node, document: &XmlDocument) -> Result<Binding> {
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

fn load_snapshot_resources(
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

fn validate_model(extension: &AddIn) -> Result<()> {
    validate_model_with(extension, &Limits::standard())
}

fn validate_model_with(extension: &AddIn, limits: &Limits) -> Result<()> {
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

fn validate_binding(binding: &Binding) -> Result<()> {
    require_nonempty("binding id", &binding.id)?;
    require_nonempty("binding type", binding.kind.as_str())?;
    require_nonempty("binding appref", &binding.app_ref)?;
    validate_extension_list(binding.extension_list.as_ref(), &[ExtKind::AddIn])
}

fn validate_store_reference(reference: &Reference) -> Result<()> {
    require_nonempty("reference id", &reference.id)?;
    require_nonempty("reference version", &reference.version)?;
    validate_extension_list(reference.extension_list.as_ref(), &[ExtKind::AddIn])
}

fn validate_task_pane(pane: &Pane) -> Result<()> {
    validate_task_pane_with(pane, &Limits::standard())
}

fn validate_task_pane_with(pane: &Pane, limits: &Limits) -> Result<()> {
    require_nonempty("dock state", pane.dock_state.as_str())?;
    require_nonempty("task-pane relationship id", &pane.relationship_id)?;
    if !pane.width.is_finite() || pane.width <= 0.0 {
        return invalid("task-pane width must be finite and positive".into());
    }
    validate_extension_list(pane.extension_list.as_ref(), &[ExtKind::TaskPane])?;
    validate_model_with(&pane.add_in, limits)?;
    validate_snapshot_resources_with(pane, limits)
}

fn validate_panes(task_panes: &Panes, limits: &Limits) -> Result<()> {
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

fn charge_authored_metadata(
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

fn add_xml_budget(total: &mut usize, bytes: usize, limits: &Limits) -> Result<()> {
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

fn add_escaped_xml_budget(total: &mut usize, value: &str, limits: &Limits) -> Result<()> {
    let bytes = value.len().checked_mul(6).ok_or(Error::Limit {
        resource: "authored web extension XML bytes",
        max: limits.xml_bytes,
        actual: usize::MAX,
    })?;
    add_xml_budget(total, bytes, limits)
}

fn add_reference_budget(total: &mut usize, reference: &Reference, limits: &Limits) -> Result<()> {
    add_xml_budget(total, 128, limits)?;
    add_escaped_xml_budget(total, &reference.id, limits)?;
    add_escaped_xml_budget(total, &reference.version, limits)?;
    if let Some(catalog) = &reference.catalog {
        add_escaped_xml_budget(total, catalog, limits)?;
    }
    if let Some(extension_list) = &reference.extension_list {
        add_xml_budget(total, extension_list.xml.len(), limits)?;
    }
    Ok(())
}

fn validate_add_in_budget(extension: &AddIn, limits: &Limits) -> Result<()> {
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

fn validate_panes_budget(task_panes: &Panes, limits: &Limits) -> Result<()> {
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

fn validate_extension_list(extension_list: Option<&ExtList>, allowed: &[ExtKind]) -> Result<()> {
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

fn validate_snapshot_resources_with(pane: &Pane, limits: &Limits) -> Result<()> {
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

fn write_store_reference(out: &mut String, element: &str, reference: &Reference) {
    out.push_str("<we:");
    out.push_str(element);
    out.push_str(" id=\"");
    escape_attr(out, &reference.id);
    out.push_str("\" version=\"");
    escape_attr(out, &reference.version);
    if let Some(store) = &reference.catalog {
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

fn effective_namespaces(scope: &NamespaceScope) -> Result<Vec<(&String, &String)>> {
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

fn retained_namespace_bytes(
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

fn escaped_attr_bytes(value: &str) -> Result<usize> {
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

fn validate_image_content_type(value: &str) -> Result<()> {
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

fn is_mime_token_byte(byte: u8) -> bool {
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

fn validate_external_uri_reference(value: &str) -> Result<()> {
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

fn require_content_type(part: &dyn Part, expected: &str) -> Result<()> {
    if part.content_type() != expected {
        Err(Error::ContentType {
            expected: expected.into(),
            actual: part.content_type().into(),
        })
    } else {
        Ok(())
    }
}

fn checked_internal_target(
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

fn enforce_count_with(label: &'static str, count: usize, limits: &Limits) -> Result<()> {
    if count > limits.items {
        limit(label, limits.items, count)
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
    Err(Error::Invalid(message))
}

fn limit<T>(resource: &'static str, max: usize, actual: usize) -> Result<T> {
    Err(Error::Limit {
        resource,
        max,
        actual,
    })
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
    namespaces: Arc<NamespaceScope>,
    declared_prefixes: HashSet<String>,
}

#[derive(Debug)]
struct NamespaceScope {
    parent: Option<Arc<NamespaceScope>>,
    local: HashMap<String, String>,
    binding_count: usize,
}

impl NamespaceScope {
    fn xml() -> Arc<Self> {
        Arc::new(Self {
            parent: None,
            local: HashMap::from([("xml".into(), "http://www.w3.org/XML/1998/namespace".into())]),
            binding_count: 1,
        })
    }

    fn get(&self, prefix: &str) -> Option<&str> {
        self.local
            .get(prefix)
            .map(String::as_str)
            .or_else(|| self.parent.as_deref().and_then(|parent| parent.get(prefix)))
    }
}

#[derive(Debug)]
struct NodeFrame {
    node: Node,
    namespaces: Arc<NamespaceScope>,
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
    string_bytes: usize,
}

impl XmlDocument {
    fn root(&self) -> Result<&Node> {
        self.root
            .as_ref()
            .ok_or_else(|| Error::Invalid("missing XML root".into()))
    }

    fn self_contained_fragment(&self, node: &Node) -> Result<String> {
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

fn parse_mce_xml(xml: &[u8], namespaces: &[&str], limits: &Limits) -> Result<XmlDocument> {
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

fn parse_xml(xml: &[u8]) -> Result<XmlDocument> {
    let limits = Limits::standard();
    if xml.len() > limits.xml_bytes {
        return limit("web extension XML bytes", limits.xml_bytes, xml.len());
    }
    parse_xml_owned(xml.to_vec(), &limits)
}

fn parse_xml_owned(xml: Vec<u8>, limits: &Limits) -> Result<XmlDocument> {
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
struct ElementEvent {
    empty: bool,
    start: usize,
    end: usize,
}

fn push_element(
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
        Error::Invalid(format!(
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
    use litchi_opc::constants::relationship_type as rt;
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
        let panes = load(&omex).unwrap().unwrap();
        assert_eq!(panes.panes.len(), 1);
        assert_eq!(panes.panes[0].add_in.reference.store, Store::Omex);
        assert!(panes.panes[0].visible);

        let registry = local_fixture_package(LOCAL_HIDDEN_TASK_PANES, LOCAL_REGISTRY_EXTENSION);
        let panes = load(&registry).unwrap().unwrap();
        assert_eq!(panes.panes[0].add_in.reference.store, Store::Registry);
        assert!(!panes.panes[0].visible);
    }

    #[test]
    fn strict_writer_is_deterministic_and_round_trips() {
        let extension = sample_extension();
        let first = write_add_in(&extension, Conformance::Strict).unwrap();
        let second = write_add_in(&extension, Conformance::Strict).unwrap();
        assert_eq!(first, second);
        assert!(
            std::str::from_utf8(&first)
                .unwrap()
                .contains(STRICT_RELATIONSHIPS_NAMESPACE)
        );
        assert_eq!(parse_add_in(&first).unwrap(), extension);
    }

    #[test]
    fn snapshot_compression_and_effect_trees_round_trip() {
        let extension = parse_add_in(LOCAL_SNAPSHOT_EFFECTS_EXTENSION).unwrap();
        let snapshot = extension.snapshot.as_ref().unwrap();
        assert_eq!(
            snapshot.compression_state,
            Some(Compression::HighQualityPrint)
        );
        assert_eq!(
            snapshot
                .effects
                .iter()
                .map(Effect::kind)
                .collect::<Vec<_>>(),
            vec![
                EffectKind::AlphaModulateFixed,
                EffectKind::Duotone,
                EffectKind::Blur,
            ]
        );
        assert!(snapshot.effects[1].xml().contains("srgbClr"));

        let written = write_add_in(&extension, Conformance::Strict).unwrap();
        let reparsed = parse_add_in(&written).unwrap();
        assert_eq!(reparsed, extension);
        let written = std::str::from_utf8(&written).unwrap();
        assert!(written.contains("cstate=\"hqprint\""));
        assert!(written.contains(STRICT_RELATIONSHIPS_NAMESPACE));
    }

    #[test]
    fn preserves_all_extension_list_sites_with_inherited_namespaces_and_mixed_content() {
        let extension = parse_add_in(LOCAL_EXTENSION_LISTS_EXTENSION).unwrap();
        let reference_extension = extension.reference.extension_list.as_ref().unwrap();
        assert_eq!(reference_extension.kind(), ExtKind::AddIn);
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
            ExtKind::DrawingMl
        );
        assert!(extension.extension_list.is_some());

        let written = write_add_in(&extension, Conformance::Strict).unwrap();
        assert_eq!(parse_add_in(&written).unwrap(), extension);

        let panes = parse_panes(LOCAL_EXTENSION_LISTS_TASK_PANES).unwrap();
        let pane_extension = panes[0].extension_list.as_ref().unwrap();
        assert_eq!(pane_extension.kind(), ExtKind::TaskPane);
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
        let loaded = load(&package).unwrap().unwrap();
        assert!(loaded.panes[0].extension_list.is_some());
        assert!(loaded.panes[0].add_in.extension_list.is_some());

        let mut stored = OpcPackage::new();
        put(&mut stored, loaded.clone(), Conformance::Strict).unwrap();
        assert_eq!(load(&stored).unwrap(), Some(loaded));
    }

    #[test]
    fn authored_extension_lists_validate_namespace_placement_and_security() {
        let web = ExtList::from_xml(
            br#"<we:extLst xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"><a:ext xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uri="urn:test"><v:data xmlns:v="urn:test">text<![CDATA[data]]></v:data></a:ext></we:extLst>"#,
        )
        .unwrap();
        assert_eq!(web.kind(), ExtKind::AddIn);
        assert_eq!(ExtList::from_xml(web.as_xml()).unwrap(), web);

        let mut extension = sample_extension();
        extension.extension_list = Some(web.clone());
        assert!(write_add_in(&extension, Conformance::Transitional).is_ok());

        let mut panes = sample_task_panes();
        panes.panes[0].extension_list = Some(web);
        assert!(write_panes(&panes, Conformance::Transitional).is_err());
        assert!(
            ExtList::from_xml(
                br#"<!DOCTYPE extLst [<!ENTITY x "boom">]><we:extLst xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11">&x;</we:extLst>"#
            )
            .is_err()
        );
        assert!(
            ExtList::from_xml(br#"<v:extLst xmlns:v="urn:not-an-office-namespace"/>"#).is_err()
        );
        assert!(
            ExtList::from_xml(
                br#"<we:extLst xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" unexpected="1"/>"#
            )
            .is_err()
        );
        assert!(
            ExtList::from_xml(
                br#"<we:extLst xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ext/></we:extLst>"#
            )
            .is_err()
        );
        assert!(
            ExtList::from_xml(
                br#"<we:extLst xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" xmlns:v="urn:test"><v:data/></we:extLst>"#
            )
            .is_err()
        );
        assert!(
            ExtList::from_xml(
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
            let effect = Effect::from_xml(xml.as_bytes()).unwrap();
            assert_eq!(effect.kind().local_name(), name);
            assert_eq!(Effect::from_xml(effect.xml().as_bytes()).unwrap(), effect);
        }

        assert!(
            Effect::from_xml(
                br#"<a:reflection xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#
            )
            .is_err()
        );
        assert!(
            Effect::from_xml(
                br#"<!DOCTYPE x><a:blur xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#
            )
            .is_err()
        );
        assert!(
            Effect::from_xml(
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
        assert!(parse_add_in(invalid_compression.as_bytes()).is_err());

        let misplaced_extension_list = format!(
            r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" xmlns:a="{DRAWINGML_NAMESPACE}" id="x"><we:reference id="a" version="1"/><we:properties/><we:bindings/><we:snapshot><a:extLst/><a:blur/></we:snapshot></we:webextension>"#
        );
        assert!(parse_add_in(misplaced_extension_list.as_bytes()).is_err());
    }

    #[test]
    fn accepts_mce_alternate_content_and_strict_relationship_attributes() {
        let xml = format!(
            r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" xmlns:r="{STRICT_RELATIONSHIPS_NAMESPACE}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" id="x"><we:reference id="a" version="1"/><mc:AlternateContent><mc:Choice Requires="we"><we:alternateReferences/></mc:Choice><mc:Fallback/></mc:AlternateContent><we:properties/><we:bindings/><we:snapshot r:embed="rId1"/></we:webextension>"#
        );
        let extension = parse_add_in(xml.as_bytes()).unwrap();
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
        assert!(parse_add_in(br#"<!DOCTYPE x><x/>"#).is_err());
        let bad_order = format!(
            r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" id="x"><we:properties/><we:reference id="a" version="1"/><we:bindings/></we:webextension>"#
        );
        assert!(parse_add_in(bad_order.as_bytes()).is_err());
        let bad_store = format!(
            r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" id="x"><we:reference id="a" version="1" storeType="Network"/><we:properties/><we:bindings/></we:webextension>"#
        );
        assert!(parse_add_in(bad_store.as_bytes()).is_err());
        let bad_width = format!(
            r#"<wetp:taskpanes xmlns:wetp="{TASK_PANES_NAMESPACE}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS_NAMESPACE}"><wetp:taskpane dockstate="right" visibility="1" width="NaN" row="0"><wetp:webextensionref r:id="rId1"/></wetp:taskpane></wetp:taskpanes>"#
        );
        assert!(parse_panes(bad_width.as_bytes()).is_err());
        let obsolete_float = format!(
            r#"<wetp:taskpanes xmlns:wetp="{TASK_PANES_NAMESPACE}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS_NAMESPACE}"><wetp:taskpane dockstate="right" visibility="1" width="320" row="0"><wetp:webextensionref r:id="rId1"/><wetp:float/></wetp:taskpane></wetp:taskpanes>"#
        );
        assert!(parse_panes(obsolete_float.as_bytes()).is_err());
    }

    #[test]
    fn enforces_input_and_list_caps() {
        assert!(parse_add_in(&vec![b' '; MAX_WEB_EXTENSION_XML_BYTES + 1]).is_err());
        let mut model = Panes::default();
        model
            .panes
            .resize_with(MAX_WEB_EXTENSION_ITEMS + 1, || Pane {
                dock_state: Dock::Right,
                visible: false,
                width: 320.0,
                row: 0,
                locked: false,
                relationship_id: "rId1".into(),
                add_in: sample_extension(),
                snapshot_resources: vec![],
                extension_list: None,
            });
        assert!(write_panes(&model, Conformance::Transitional).is_err());
        let mut extension = sample_extension();
        extension.id = "x".repeat(MAX_WEB_EXTENSION_XML_BYTES);
        assert!(write_add_in(&extension, Conformance::Transitional).is_err());

        let mut excessive_nodes = format!(
            r#"<we:extLst xmlns:we="{WEB_EXTENSION_NAMESPACE}" xmlns:a="{DRAWINGML_NAMESPACE}" xmlns:v="urn:test"><a:ext uri="urn:test"><v:data>"#
        );
        excessive_nodes.push_str(&"<v:n/>".repeat(MAX_WEB_EXTENSION_XML_NODES));
        excessive_nodes.push_str("</v:data></a:ext></we:extLst>");
        assert!(ExtList::from_xml(excessive_nodes.as_bytes()).is_err());
    }

    #[test]
    fn rejects_external_wrong_content_type_and_dangling_package_graphs() {
        let external = synthetic_package(true, ADD_IN_CONTENT_TYPE, "rId1");
        assert!(load(&external).is_err());

        let wrong_type = synthetic_package(false, "application/xml", "rId1");
        assert!(matches!(load(&wrong_type), Err(Error::ContentType { .. })));

        let dangling = synthetic_package(false, ADD_IN_CONTENT_TYPE, "missing");
        assert!(load(&dangling).is_err());
    }

    #[test]
    fn package_crud_round_trips_embedded_and_linked_snapshots() {
        let mut package = OpcPackage::new();
        let authored = sample_task_panes();
        put(&mut package, authored.clone(), Conformance::Transitional).unwrap();
        assert_eq!(load(&package).unwrap(), Some(authored.clone()));

        let task_panes_name = package
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
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
        replacement.panes[0].add_in.snapshot = None;
        replacement.panes[0].snapshot_resources.clear();
        replacement.panes[0].visible = false;
        put(&mut package, replacement.clone(), Conformance::Strict).unwrap();
        assert_eq!(load(&package).unwrap(), Some(replacement));
        assert!(
            package
                .get_part(&PackURI::new("/media/web-extension-snapshot.png").unwrap())
                .is_err()
        );

        assert!(remove(&mut package).unwrap());
        assert!(load(&package).unwrap().is_none());
        assert!(!remove(&mut package).unwrap());
        assert_eq!(package.part_count(), 0);
    }

    #[test]
    fn byte_identical_put_is_a_signature_preserving_no_op() {
        let mut package = OpcPackage::new();
        put(&mut package, sample_task_panes(), Conformance::Transitional).unwrap();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        assert!(package.is_signed());

        let loaded = load(&package).unwrap().unwrap();
        put(&mut package, loaded, Conformance::Transitional).unwrap();

        assert!(package.is_signed());
    }

    #[test]
    fn changed_shared_task_pane_part_is_rejected_without_mutation() {
        let mut package = OpcPackage::new();
        put(&mut package, sample_task_panes(), Conformance::Transitional).unwrap();
        let task_panes_name = package
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
            .unwrap()
            .target_partname()
            .unwrap();
        let target = task_panes_name.as_str().trim_start_matches('/').to_owned();
        package.rels_mut().add_relationship(
            "urn:litchi:test:shared-task-panes".into(),
            target,
            "rIdSharedTaskPanes".into(),
            false,
        );
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        let before = package.get_part(&task_panes_name).unwrap().blob().to_vec();
        let mut changed = load(&package).unwrap().unwrap();
        changed.panes[0].visible = false;

        assert!(put(&mut package, changed, Conformance::Transitional).is_err());
        assert_eq!(package.get_part(&task_panes_name).unwrap().blob(), before);
        assert!(package.is_signed());
        assert!(load(&package).unwrap().unwrap().panes[0].visible);
    }

    #[test]
    fn shared_task_pane_ingress_protects_descendant_add_in() {
        let mut package = OpcPackage::new();
        put(&mut package, sample_task_panes(), Conformance::Transitional).unwrap();
        let task_panes_name = add_shared_task_pane_ingress(&mut package);
        let extension_name = package
            .get_part(&task_panes_name)
            .unwrap()
            .rels()
            .get("rId1")
            .unwrap()
            .target_partname()
            .unwrap();
        let before = package.get_part(&extension_name).unwrap().blob().to_vec();
        let mut changed = load(&package).unwrap().unwrap();
        assert!(
            changed
                .edit(0usize, |pane| {
                    pane.add_in_mut().set_frozen(false);
                    Ok(())
                })
                .unwrap()
        );

        assert!(put(&mut package, changed, Conformance::Transitional).is_err());
        assert_eq!(package.get_part(&extension_name).unwrap().blob(), before);
        assert!(
            load(&package)
                .unwrap()
                .unwrap()
                .get(0usize)
                .unwrap()
                .add_in()
                .is_frozen()
        );
    }

    #[test]
    fn shared_task_pane_ingress_protects_descendant_image() {
        let mut package = OpcPackage::new();
        put(&mut package, sample_task_panes(), Conformance::Transitional).unwrap();
        add_shared_task_pane_ingress(&mut package);
        let image_name = PackURI::new("/media/web-extension-snapshot.png").unwrap();
        let before = package.get_part(&image_name).unwrap().blob().to_vec();
        let mut changed = load(&package).unwrap().unwrap();
        assert!(
            changed
                .edit(0usize, |pane| {
                    pane.set_image(
                        "/media/web-extension-snapshot.png",
                        "image/png",
                        Arc::new(vec![1, 2, 3, 4]),
                    )?;
                    Ok(())
                })
                .unwrap()
        );

        assert!(put(&mut package, changed, Conformance::Transitional).is_err());
        assert_eq!(package.get_part(&image_name).unwrap().blob(), before);
    }

    #[test]
    fn internal_web_relationship_targets_reject_queries_and_fragments() {
        let mut root = OpcPackage::new();
        put(&mut root, sample_task_panes(), Conformance::Transitional).unwrap();
        let root_id = root
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
            .unwrap()
            .r_id()
            .to_owned();
        root.rels_mut().remove(&root_id);
        root.rels_mut().add_relationship(
            TASK_PANES_RELATIONSHIP.into(),
            "webextensions/taskpanes.xml?version=1".into(),
            root_id,
            false,
        );
        assert!(matches!(load(&root), Err(Error::Relationship(_))));

        let mut add_in = OpcPackage::new();
        put(&mut add_in, sample_task_panes(), Conformance::Transitional).unwrap();
        let task_name = add_in
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
            .unwrap()
            .target_partname()
            .unwrap();
        let task_part = add_in.get_part_mut(&task_name).unwrap();
        task_part.rels_mut().remove("rId1");
        task_part.rels_mut().add_relationship(
            ADD_IN_RELATIONSHIP.into(),
            "webextension1.xml#instance".into(),
            "rId1".into(),
            false,
        );
        assert!(matches!(load(&add_in), Err(Error::Relationship(_))));

        let mut image = OpcPackage::new();
        put(&mut image, sample_task_panes(), Conformance::Transitional).unwrap();
        let task_name = image
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
            .unwrap()
            .target_partname()
            .unwrap();
        let add_in_name = image
            .get_part(&task_name)
            .unwrap()
            .rels()
            .get("rId1")
            .unwrap()
            .target_partname()
            .unwrap();
        let add_in_part = image.get_part_mut(&add_in_name).unwrap();
        add_in_part.rels_mut().remove("rIdSnapshot");
        add_in_part.rels_mut().add_relationship(
            IMAGE_RELATIONSHIP_TYPE.into(),
            "../media/web-extension-snapshot.png?size=large".into(),
            "rIdSnapshot".into(),
            false,
        );
        assert!(matches!(load(&image), Err(Error::Relationship(_))));
    }

    #[test]
    fn case_equivalent_existing_parts_are_rejected_before_put_mutates() {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/Data.bin").unwrap(),
            "application/octet-stream".into(),
            vec![1],
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/CUSTOM/data.bin").unwrap(),
            "application/octet-stream".into(),
            vec![2],
        )));
        let before = package.part_count();
        assert!(put(&mut package, sample_task_panes(), Conformance::Transitional,).is_err());
        assert_eq!(package.part_count(), before);
        assert_eq!(package.rels().len(), 0);
    }

    #[test]
    fn package_store_rejects_resource_mismatches_without_mutation() {
        let mut package = OpcPackage::new();
        let mut malformed = sample_task_panes();
        malformed.panes[0].snapshot_resources.pop();
        assert!(put(&mut package, malformed, Conformance::Transitional,).is_err());
        assert_eq!(package.part_count(), 0);
        assert_eq!(package.rels().iter().count(), 0);

        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/media/web-extension-snapshot.png").unwrap(),
            "image/png".into(),
            vec![9, 9, 9],
        )));
        assert!(put(&mut package, sample_task_panes(), Conformance::Transitional,).is_err());
        assert_eq!(package.part_count(), 1);
        assert_eq!(
            package
                .get_part(&PackURI::new("/media/web-extension-snapshot.png").unwrap())
                .unwrap()
                .blob(),
            &[9, 9, 9]
        );
    }

    #[test]
    fn semantic_facade_authors_selects_and_removes_without_raw_ids() {
        let reference = Reference::new("wa1", "1.0.0.0", Store::Omex).unwrap();
        let binding = Binding::new("binding-1", BindingKind::Matrix, "app-ref").unwrap();
        let add_in = AddIn::new("add-in-1", reference)
            .unwrap()
            .bind(binding)
            .unwrap();
        let bytes = Arc::new(vec![0x89, b'P', b'N', b'G']);
        let effect = Effect::from_xml(
            br#"<a:blur xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" rad="1000"/>"#,
        )
        .unwrap();
        let pane = Pane::new(add_in)
            .show(false)
            .dock(Dock::Left)
            .unwrap()
            .width(420.0)
            .unwrap()
            .embed("/media/add-in-preview.png", "image/png", Arc::clone(&bytes))
            .unwrap()
            .linked("https://example.invalid/inert.png#preview")
            .unwrap()
            .compress(Compression::Print)
            .effect(effect)
            .unwrap();

        let mut panes = Panes::new();
        panes.push(pane).unwrap();
        assert_eq!(panes.get("add-in-1").unwrap().dock_kind(), &Dock::Left);
        assert!(!panes.get(0usize).unwrap().visible());
        assert!(panes.get(1usize).is_none());
        let image = panes.get("add-in-1").unwrap().image().unwrap();
        assert!(Arc::ptr_eq(&bytes, &image.shared()));
        assert_eq!(image.content_type(), "image/png");
        assert_eq!(
            panes.get("add-in-1").unwrap().link().unwrap().external(),
            Some("https://example.invalid/inert.png#preview")
        );
        assert!(panes.remove(99usize).is_none());
        assert_eq!(panes.remove("add-in-1").unwrap().add_in().id(), "add-in-1");
        assert!(panes.is_empty());
    }

    #[test]
    fn panes_push_rekeys_colliding_hidden_relationship_ids() {
        let first_add_in = AddIn::new(
            "add-in-1",
            Reference::new("ref-1", "1", Store::Omex).unwrap(),
        )
        .unwrap();
        let second_add_in = AddIn::new(
            "add-in-2",
            Reference::new("ref-2", "1", Store::Registry).unwrap(),
        )
        .unwrap();
        let mut first_source = Panes::new();
        first_source.push(Pane::new(first_add_in)).unwrap();
        let mut second_source = Panes::new();
        second_source.push(Pane::new(second_add_in)).unwrap();
        let first = first_source.remove(0usize).unwrap();
        let second = second_source.remove(0usize).unwrap();
        let mut panes = Panes::new();
        panes.push(first).unwrap();
        panes.push(second).unwrap();

        assert_eq!(panes.len(), 2);
        assert_ne!(
            panes.panes[0].relationship_id,
            panes.panes[1].relationship_id
        );
        assert!(!panes.panes[1].relationship_id.is_empty());
        let mut package = OpcPackage::new();
        put(&mut package, panes, Conformance::Transitional).unwrap();
        assert_eq!(load(&package).unwrap().unwrap().len(), 2);
    }

    #[test]
    fn panes_push_canonicalizes_equivalent_resources_within_one_pane() {
        let bytes = Arc::new(vec![1, 2, 3]);
        let pane = Pane::new(
            AddIn::new(
                "add-in-1",
                Reference::new("ref-1", "1", Store::Omex).unwrap(),
            )
            .unwrap(),
        )
        .embed("/media/Preview.png", "image/png", Arc::clone(&bytes))
        .unwrap()
        .linked_image("/MEDIA/preview.png", "image/png", Arc::clone(&bytes))
        .unwrap();

        let mut panes = Panes::new();
        panes.push(pane).unwrap();

        let pane = panes.get("add-in-1").unwrap();
        let embedded = pane.image().unwrap();
        let linked = pane.link().unwrap().internal().unwrap();
        assert_eq!(embedded.name(), linked.name());
        assert_eq!(embedded.name().as_str(), "/media/Preview.png");
        assert!(Arc::ptr_eq(&embedded.shared(), &linked.shared()));
        let mut package = OpcPackage::new();
        put(&mut package, panes, Conformance::Transitional).unwrap();
    }

    #[test]
    fn panes_push_rejects_conflicting_resources_within_one_pane_atomically() {
        let pane = Pane::new(
            AddIn::new(
                "add-in-1",
                Reference::new("ref-1", "1", Store::Omex).unwrap(),
            )
            .unwrap(),
        )
        .embed("/media/Preview.png", "image/png", vec![1, 2, 3])
        .unwrap()
        .linked_image("/MEDIA/preview.png", "image/png", vec![4, 5, 6])
        .unwrap();

        let mut panes = Panes::new();
        assert!(panes.push(pane).is_err());
        assert!(panes.is_empty());
    }

    #[test]
    fn panes_edit_is_checked_canonical_and_transactional() {
        let first = Pane::new(
            AddIn::new(
                "add-in-1",
                Reference::new("ref-1", "1", Store::Omex).unwrap(),
            )
            .unwrap(),
        )
        .embed("/media/Preview.png", "image/png", vec![1, 2, 3])
        .unwrap();
        let second = Pane::new(
            AddIn::new(
                "add-in-2",
                Reference::new("ref-2", "1", Store::Registry).unwrap(),
            )
            .unwrap(),
        )
        .embed("/media/other.png", "image/png", vec![4, 5, 6])
        .unwrap();
        let mut panes = Panes::new();
        panes.push(first).unwrap().push(second).unwrap();
        let before = panes.get("add-in-2").unwrap().clone();

        assert!(
            panes
                .edit("add-in-2", |pane| {
                    pane.set_visible(false);
                    pane.set_image("/MEDIA/preview.png", "image/png", vec![9])?;
                    Ok(())
                })
                .is_err()
        );
        assert_eq!(panes.get("add-in-2"), Some(&before));

        let shared = panes.get("add-in-1").unwrap().image().unwrap().shared();
        assert!(
            panes
                .edit("add-in-2", |pane| {
                    pane.set_visible(false);
                    pane.set_image("/MEDIA/preview.png", "image/png", Arc::clone(&shared))?;
                    Ok(())
                })
                .unwrap()
        );
        let edited = panes.get("add-in-2").unwrap();
        assert!(!edited.visible());
        assert_eq!(
            edited.image().unwrap().name().as_str(),
            "/media/Preview.png"
        );
        assert!(Arc::ptr_eq(&shared, &edited.image().unwrap().shared()));

        let mut invoked = false;
        assert!(
            !panes
                .edit(99usize, |_| {
                    invoked = true;
                    Ok(())
                })
                .unwrap()
        );
        assert!(!invoked);
    }

    #[test]
    fn conflicting_case_equivalent_snapshot_resources_are_rejected_atomically() {
        let first = Pane::new(
            AddIn::new(
                "add-in-1",
                Reference::new("ref-1", "1", Store::Omex).unwrap(),
            )
            .unwrap(),
        )
        .embed("/media/Preview.png", "image/png", vec![1, 2, 3])
        .unwrap();
        let second = Pane::new(
            AddIn::new(
                "add-in-2",
                Reference::new("ref-2", "1", Store::Registry).unwrap(),
            )
            .unwrap(),
        )
        .embed("/MEDIA/preview.png", "image/png", vec![4, 5, 6])
        .unwrap();
        let mut first_source = Panes::new();
        first_source.push(first).unwrap();
        let mut second_source = Panes::new();
        second_source.push(second).unwrap();
        let first = first_source.remove(0usize).unwrap();
        let second = second_source.remove(0usize).unwrap();

        let mut panes = Panes::new();
        panes.push(first).unwrap();
        assert!(panes.push(second).is_err());
        assert_eq!(panes.len(), 1);
        assert_eq!(
            panes.get("add-in-1").unwrap().image().unwrap().bytes(),
            &[1, 2, 3]
        );
    }

    #[test]
    fn absent_snapshot_metadata_no_ops_do_not_create_a_snapshot() {
        let add_in = AddIn::new(
            "add-in-1",
            Reference::new("ref-1", "1", Store::Omex).unwrap(),
        )
        .unwrap();
        let mut pane = Pane::new(add_in);
        assert!(pane.add_in().snapshot().is_none());

        assert!(!pane.clear_compression());
        assert!(pane.add_in().snapshot().is_none());

        let effect = Effect::from_xml(
            br#"<a:blur xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
        )
        .unwrap();
        assert!(pane.replace_effect(0, effect).unwrap().is_none());
        assert!(pane.add_in().snapshot().is_none());
    }

    #[test]
    fn semantic_snapshot_authoring_rejects_bad_mime_and_uri_reference() {
        let reference = Reference::new("wa1", "1", Store::Registry).unwrap();
        let add_in = AddIn::new("add-in-1", reference).unwrap();
        assert!(
            Pane::new(add_in.clone())
                .embed("/media/a.bin", "image", vec![1, 2, 3])
                .is_err()
        );
        assert!(
            Pane::new(add_in.clone())
                .embed("/media/a.bin", "image/png; charset=binary", vec![1])
                .is_err()
        );
        assert!(Pane::new(add_in.clone()).linked("bad target").is_err());
        assert!(
            Pane::new(add_in)
                .linked("https://example.invalid/%GG")
                .is_err()
        );
    }

    #[test]
    fn internal_linked_image_is_typed_and_round_trips() {
        let reference = Reference::new("wa1", "1", Store::Omex).unwrap();
        let add_in = AddIn::new("add-in-1", reference).unwrap();
        let pane = Pane::new(add_in)
            .linked_image(
                "/media/linked-preview.png",
                "image/png",
                Arc::new(vec![1, 2, 3, 4]),
            )
            .unwrap();
        assert_eq!(
            pane.link().unwrap().internal().unwrap().bytes(),
            &[1, 2, 3, 4]
        );
        let mut panes = Panes::new();
        panes.push(pane).unwrap();
        let mut package = OpcPackage::new();
        put(&mut package, panes, Conformance::Transitional).unwrap();
        let mut loaded = load(&package).unwrap().unwrap();
        assert_eq!(
            loaded
                .get(0usize)
                .unwrap()
                .link()
                .unwrap()
                .internal()
                .unwrap()
                .bytes(),
            &[1, 2, 3, 4]
        );
        assert!(
            loaded
                .edit(0usize, |pane| {
                    pane.set_external_link("https://example.invalid/inert.png")?;
                    Ok(())
                })
                .unwrap()
        );
        assert_eq!(
            loaded.get(0usize).unwrap().link().unwrap().external(),
            Some("https://example.invalid/inert.png")
        );
    }

    #[test]
    fn checked_update_crud_covers_collections_metadata_and_all_ext_sites() {
        let add_ext = ExtList::from_xml(
            format!(r#"<we:extLst xmlns:we="{WEB_EXTENSION_NAMESPACE}"/>"#).as_bytes(),
        )
        .unwrap();
        let pane_ext = ExtList::from_xml(
            format!(r#"<wetp:extLst xmlns:wetp="{TASK_PANES_NAMESPACE}"/>"#).as_bytes(),
        )
        .unwrap();
        let drawing_ext =
            ExtList::from_xml(format!(r#"<a:extLst xmlns:a="{DRAWINGML_NAMESPACE}"/>"#).as_bytes())
                .unwrap();

        let reference = Reference::new("primary", "1", Store::Omex).unwrap();
        let mut add_in = AddIn::new("add-in-1", reference).unwrap();
        add_in
            .set_reference(Reference::new("primary-2", "2", Store::Registry).unwrap())
            .unwrap();
        add_in
            .push_reference(Reference::new("alternate", "1", Store::FileSystem).unwrap())
            .unwrap();
        add_in
            .upsert_reference(Reference::new("alternate", "2", Store::Registry).unwrap())
            .unwrap();
        assert_eq!(
            add_in.alternate_reference("alternate").unwrap().version(),
            "2"
        );
        assert!(add_in.remove_reference(9usize).is_none());

        add_in
            .push_property(Property::new("mode", "old").unwrap())
            .unwrap();
        add_in
            .upsert_property(Property::new("mode", "new").unwrap())
            .unwrap();
        assert_eq!(add_in.property("mode").unwrap().value(), "new");
        assert!(add_in.remove_property(4usize).is_none());

        add_in
            .push_binding(Binding::new("binding", BindingKind::Matrix, "app-ref").unwrap())
            .unwrap();
        add_in
            .upsert_binding(Binding::new("binding", BindingKind::Table, "app-ref-2").unwrap())
            .unwrap();
        assert_eq!(
            add_in.binding("binding").unwrap().kind(),
            &BindingKind::Table
        );
        assert!(add_in.remove_binding(7usize).is_none());

        add_in.reference_mut().set_ext(add_ext.clone()).unwrap();
        assert!(add_in.reference_mut().clear_ext().is_some());
        add_in.reference_mut().set_ext(add_ext.clone()).unwrap();
        add_in
            .binding_mut("binding")
            .unwrap()
            .set_ext(add_ext.clone())
            .unwrap();
        add_in.set_ext(add_ext.clone()).unwrap();
        assert!(add_in.set_ext(pane_ext.clone()).is_err());

        let mut pane = Pane::new(add_in);
        pane.set_visible(false).set_row(3).set_locked(true);
        pane.set_width(480.0)
            .unwrap()
            .set_dock(Dock::Bottom)
            .unwrap();
        pane.snapshot_mut().set_ext(drawing_ext).unwrap();
        pane.set_ext(pane_ext).unwrap();
        assert!(!pane.visible());
        assert_eq!(pane.row(), 3);
        assert!(pane.locked());
        assert_eq!(pane.pane_width(), 480.0);
        assert_eq!(pane.dock_kind(), &Dock::Bottom);
        assert!(pane.add_in().reference().ext().is_some());
        assert!(pane.add_in().binding("binding").unwrap().ext().is_some());
        assert!(pane.add_in().ext().is_some());
        assert!(pane.add_in().snapshot().unwrap().ext().is_some());
        assert!(pane.ext().is_some());

        let mut panes = Panes::new();
        panes.push(pane).unwrap();
        let mut package = OpcPackage::new();
        put(&mut package, panes, Conformance::Transitional).unwrap();
        assert!(load(&package).unwrap().is_some());
    }

    #[test]
    fn package_graph_limits_bound_index_allocation_relationships_and_deletion() {
        let mut authored = OpcPackage::new();
        let no_allocations = Limits {
            part_allocations: 0,
            ..Limits::standard()
        };
        assert!(
            put_with(
                &mut authored,
                sample_task_panes(),
                Conformance::Transitional,
                &no_allocations,
            )
            .is_err()
        );
        assert_eq!(authored.part_count(), 0);

        put(
            &mut authored,
            sample_task_panes(),
            Conformance::Transitional,
        )
        .unwrap();
        let no_parts = Limits {
            package_parts: 0,
            ..Limits::standard()
        };
        assert!(matches!(
            load_with(&authored, &no_parts),
            Err(Error::Limit { .. })
        ));
        let no_relationships = Limits {
            package_relationships: 0,
            ..Limits::standard()
        };
        assert!(matches!(
            load_with(&authored, &no_relationships),
            Err(Error::Limit { .. })
        ));
        let no_deletions = Limits {
            part_deletions: 0,
            ..Limits::standard()
        };
        assert!(remove_with(&mut authored, &no_deletions).is_err());
        assert!(load(&authored).unwrap().is_some());
    }

    #[test]
    fn absent_feature_skips_the_bounded_full_graph_index() {
        let sentinel_name = PackURI::new("/custom/sentinel.bin").unwrap();
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            sentinel_name.clone(),
            "application/octet-stream".into(),
            vec![1, 2, 3],
        )));
        let limits = Limits {
            package_parts: 0,
            ..Limits::standard()
        };

        assert!(load_with(&package, &limits).unwrap().is_none());
        assert!(!remove_with(&mut package, &limits).unwrap());
        assert_eq!(package.get_part(&sentinel_name).unwrap().blob(), &[1, 2, 3]);
    }

    #[test]
    fn part_name_allocation_probes_are_operation_wide() {
        let occupied_name = PackURI::new("/webextensions/taskpanes.xml").unwrap();
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            occupied_name.clone(),
            "application/octet-stream".into(),
            vec![7],
        )));
        let mut panes = Panes::new();
        panes
            .push(Pane::new(
                AddIn::new(
                    "add-in-1",
                    Reference::new("ref-1", "1", Store::Omex).unwrap(),
                )
                .unwrap(),
            ))
            .unwrap();
        let limits = Limits {
            part_allocations: 2,
            ..Limits::standard()
        };

        assert!(matches!(
            put_with(&mut package, panes, Conformance::Transitional, &limits,),
            Err(Error::Limit { .. })
        ));
        assert_eq!(package.part_count(), 1);
        assert_eq!(package.get_part(&occupied_name).unwrap().blob(), &[7]);
    }

    #[test]
    fn aggregate_xml_and_retained_string_budgets_bound_put_and_load() {
        let panes = sample_task_panes();
        let task_xml = write_panes(&panes, Conformance::Transitional).unwrap();
        let add_in_xml = write_add_in(
            panes.get(0usize).unwrap().add_in(),
            Conformance::Transitional,
        )
        .unwrap();
        let combined_xml = task_xml.len().checked_add(add_in_xml.len()).unwrap();

        let xml_tight = Limits {
            total_xml_bytes: combined_xml - 1,
            ..Limits::standard()
        };
        let mut rejected = OpcPackage::new();
        assert!(matches!(
            put_with(
                &mut rejected,
                panes.clone(),
                Conformance::Transitional,
                &xml_tight,
            ),
            Err(Error::Limit { .. })
        ));
        assert_eq!(rejected.part_count(), 0);

        let strings_tight = Limits {
            total_string_bytes: combined_xml - 1,
            ..Limits::standard()
        };
        assert!(matches!(
            put_with(
                &mut rejected,
                panes.clone(),
                Conformance::Transitional,
                &strings_tight,
            ),
            Err(Error::Limit { .. })
        ));
        assert_eq!(rejected.part_count(), 0);

        let mut stored = OpcPackage::new();
        put(&mut stored, panes, Conformance::Transitional).unwrap();
        assert!(matches!(
            load_with(&stored, &xml_tight),
            Err(Error::Limit { .. })
        ));
        assert!(matches!(
            load_with(&stored, &strings_tight),
            Err(Error::Limit { .. })
        ));
    }

    #[test]
    fn inherited_namespace_expansion_is_charged_before_fragment_retention() {
        let mut declarations = String::new();
        for index in 0..32 {
            declarations.push_str(&format!(
                " xmlns:v{index}=\"urn:litchi:{}\"",
                "n".repeat(128)
            ));
        }
        let mut bindings = String::new();
        for index in 0..32 {
            bindings.push_str(&format!(
                r#"<we:binding id="b{index}" type="table" appref="a{index}"><we:extLst><a:ext uri="urn:e"><v0:data/></a:ext></we:extLst></we:binding>"#
            ));
        }
        let xml = format!(
            r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" xmlns:a="{DRAWINGML_NAMESPACE}"{declarations} id="x"><we:reference id="r" version="1"/><we:properties/><we:bindings>{bindings}</we:bindings></we:webextension>"#
        );
        let limits = Limits {
            string_bytes: 32 * 1024,
            ..Limits::standard()
        };

        assert!(matches!(
            parse_add_in_with(xml.as_bytes(), &limits),
            Err(Error::Limit { .. })
        ));
        assert_eq!(parse_add_in(xml.as_bytes()).unwrap().bindings().len(), 32);
    }

    #[test]
    fn package_and_authored_relationship_metadata_share_the_string_budget() {
        let mut stored = local_fixture_package(LOCAL_VISIBLE_TASK_PANES, LOCAL_OMEX_EXTENSION);
        stored.rels_mut().add_relationship(
            "urn:litchi:test:opaque".into(),
            format!("https://example.invalid/{}", "x".repeat(16 * 1024)),
            "rIdOpaque".into(),
            true,
        );
        let tight = Limits {
            total_string_bytes: 32 * 1024,
            ..Limits::standard()
        };
        assert!(matches!(
            load_with(&stored, &tight),
            Err(Error::Limit { .. })
        ));

        let add_in = AddIn::new(
            "add-in-1",
            Reference::new("ref-1", "1", Store::Omex).unwrap(),
        )
        .unwrap();
        let pane = Pane::new(add_in)
            .linked(format!("https://example.invalid/{}", "y".repeat(16 * 1024)))
            .unwrap();
        let mut panes = Panes::new();
        panes.push(pane).unwrap();
        let mut package = OpcPackage::new();
        assert!(matches!(
            put_with(&mut package, panes, Conformance::Transitional, &tight,),
            Err(Error::Limit { .. })
        ));
        assert_eq!(package.part_count(), 0);
    }

    #[test]
    fn explicit_limits_bound_both_put_and_load() {
        let reference = Reference::new("wa1", "1", Store::Omex).unwrap();
        let add_in = AddIn::new("add-in-1", reference).unwrap();
        let mut panes = Panes::new();
        panes.push(Pane::new(add_in)).unwrap();
        let tight = Limits {
            xml_bytes: 128,
            ..Limits::standard()
        };

        let mut rejected = OpcPackage::new();
        assert!(
            put_with(
                &mut rejected,
                panes.clone(),
                Conformance::Transitional,
                &tight,
            )
            .is_err()
        );
        assert_eq!(rejected.part_count(), 0);
        assert_eq!(rejected.rels().len(), 0);

        let mut stored = OpcPackage::new();
        put(&mut stored, panes, Conformance::Transitional).unwrap();
        assert!(load_with(&stored, &tight).is_err());
    }

    fn synthetic_package(
        external_extension: bool,
        extension_content_type: &str,
        pane_relationship_id: &str,
    ) -> OpcPackage {
        let task_panes_xml = format!(
            r#"<wetp:taskpanes xmlns:wetp="{TASK_PANES_NAMESPACE}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS_NAMESPACE}"><wetp:taskpane dockstate="right" visibility="0" width="320" row="0"><wetp:webextensionref r:id="{pane_relationship_id}"/></wetp:taskpane></wetp:taskpanes>"#
        );
        let extension_xml = write_add_in(&sample_extension(), Conformance::Transitional).unwrap();
        let mut package = OpcPackage::new();
        package.rels_mut().add_relationship(
            TASK_PANES_RELATIONSHIP.into(),
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
            ADD_IN_RELATIONSHIP.into(),
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

    fn add_shared_task_pane_ingress(package: &mut OpcPackage) -> PackURI {
        let task_panes_name = package
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
            .unwrap()
            .target_partname()
            .unwrap();
        package.rels_mut().add_relationship(
            "urn:litchi:test:shared-task-panes".into(),
            task_panes_name.as_str().trim_start_matches('/').to_owned(),
            "rIdSharedTaskPanes".into(),
            false,
        );
        task_panes_name
    }

    fn local_fixture_package(task_panes_xml: &[u8], extension_xml: &[u8]) -> OpcPackage {
        let mut package = OpcPackage::new();
        package.rels_mut().add_relationship(
            TASK_PANES_RELATIONSHIP.into(),
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
            ADD_IN_RELATIONSHIP.into(),
            "webextension1.xml".into(),
            "rId1".into(),
            false,
        );
        package.add_part(Box::new(task_panes_part));
        package.add_part(Box::new(XmlPart::new(
            PackURI::new("/webextensions/webextension1.xml").unwrap(),
            ADD_IN_CONTENT_TYPE.into(),
            extension_xml.to_vec(),
        )));
        package
    }

    fn sample_extension() -> AddIn {
        AddIn {
            id: "{00000000-0000-0000-0000-000000000001}".into(),
            frozen: true,
            reference: Reference {
                id: "wa1".into(),
                version: "1.0.0.0".into(),
                catalog: Some("en-us".into()),
                store: Store::Omex,
                extension_list: None,
            },
            alternate_references: vec![],
            properties: vec![Property {
                name: "Office.AutoShowTaskpaneWithDocument".into(),
                value: "false".into(),
            }],
            bindings: vec![Binding {
                id: "binding-1".into(),
                kind: BindingKind::Matrix,
                app_ref: "app-ref".into(),
                extension_list: None,
            }],
            snapshot: Some(Snapshot::default()),
            extension_list: None,
        }
    }

    fn sample_task_panes() -> Panes {
        let mut extension = sample_extension();
        extension.snapshot = Some(Snapshot {
            embedded_relationship_id: Some("rIdSnapshot".into()),
            linked_relationship_id: Some("rIdLinked".into()),
            compression_state: Some(Compression::HighQualityPrint),
            effects: vec![
                Effect::from_xml(
                    br#"<a:alphaModFix xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" amt="50000"/>"#,
                )
                .unwrap(),
            ],
            extension_list: None,
        });
        Panes {
            panes: vec![Pane {
                dock_state: Dock::Right,
                visible: true,
                width: 320.0,
                row: 0,
                locked: false,
                relationship_id: "rId1".into(),
                add_in: extension,
                snapshot_resources: vec![
                    SnapshotResource {
                        relationship_id: "rIdSnapshot".into(),
                        target: SnapshotTarget::Internal {
                            part_name: PackURI::new("/media/web-extension-snapshot.png").unwrap(),
                            content_type: "image/png".into(),
                            data: Arc::new(vec![0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a]),
                        },
                    },
                    SnapshotResource {
                        relationship_id: "rIdLinked".into(),
                        target: SnapshotTarget::External {
                            target: "https://example.invalid/inert-snapshot.png".into(),
                        },
                    },
                ],
                extension_list: None,
            }],
        }
    }
}
