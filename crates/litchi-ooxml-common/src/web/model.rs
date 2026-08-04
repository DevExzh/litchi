//! Validated MS-OWEXML data models and resource-budget invariants.

use super::codec::*;
use super::package::*;
use super::*;

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

    pub(super) fn parse(value: &str) -> Result<Self> {
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
pub(super) const MAX_WEB_EXTENSION_XML_BYTES: usize = STANDARD_XML_BYTES;
#[cfg(test)]
pub(super) const MAX_WEB_EXTENSION_XML_NODES: usize = STANDARD_NODES;
pub(super) const MAX_WEB_EXTENSION_ITEMS: usize = STANDARD_ITEMS;
pub(super) const MAX_WEB_EXTENSION_SNAPSHOT_BYTES: usize = STANDARD_IMAGE_BYTES;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub(super) fn relationships_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_RELATIONSHIPS_NAMESPACE,
            Self::Strict => STRICT_RELATIONSHIPS_NAMESPACE,
        }
    }

    pub(super) fn image_relationship_type(self) -> &'static str {
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
    /// File-system provider. Author references with [`Reference::file`].
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

    pub(super) fn parse(value: &str) -> Result<Self> {
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
    pub(super) id: String,
    pub(super) version: String,
    pub(super) location: Option<String>,
    pub(super) store: Store,
    pub(super) extension_list: Option<ExtList>,
}

impl Reference {
    /// Create a validated reference for a catalog-backed provider.
    ///
    /// File-system references require a location and must be created with
    /// [`Self::file`]. Keeping the location in the constructor prevents the
    /// safe model from representing the store-less form rejected by Office.
    pub fn new(id: impl Into<String>, version: impl Into<String>, store: Store) -> Result<Self> {
        if store == Store::FileSystem {
            return invalid(
                "FileSystem references require Reference::file(id, version, location)".into(),
            );
        }
        let value = Self {
            id: id.into(),
            version: version.into(),
            location: None,
            store,
            extension_list: None,
        };
        validate_store_reference(&value)?;
        Ok(value)
    }

    /// Create a validated file-system reference with its required location.
    pub fn file(
        id: impl Into<String>,
        version: impl Into<String>,
        location: impl Into<String>,
    ) -> Result<Self> {
        let value = Self {
            id: id.into(),
            version: version.into(),
            location: Some(location.into()),
            store: Store::FileSystem,
            extension_list: None,
        };
        validate_store_reference(&value)?;
        Ok(value)
    }

    /// Add or replace the optional provider-specific location.
    pub fn location(mut self, value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        require_nonempty("reference location", &value)?;
        self.location = Some(value);
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
    pub fn location_name(&self) -> Option<&str> {
        self.location.as_deref()
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
    pub(super) name: String,
    pub(super) value: String,
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

    pub(super) fn parse(value: &str) -> Result<Self> {
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
    pub(super) id: String,
    pub(super) kind: BindingKind,
    pub(super) app_ref: String,
    pub(super) extension_list: Option<ExtList>,
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

    pub(super) fn from_namespace(namespace: &str) -> Result<Self> {
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
    pub(super) kind: ExtKind,
    pub(super) xml: String,
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

    pub(super) fn from_node(node: &Node, document: &XmlDocument) -> Result<Self> {
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

    pub(super) fn parse(value: &str) -> Result<Self> {
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

    pub(super) fn parse(local_name: &str) -> Result<Self> {
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
    pub(super) kind: EffectKind,
    pub(super) xml: String,
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

    pub(super) fn from_node(node: &Node) -> Result<Self> {
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
    pub(super) embedded_relationship_id: Option<String>,
    pub(super) linked_relationship_id: Option<String>,
    pub(super) compression_state: Option<Compression>,
    pub(super) effects: Vec<Effect>,
    pub(super) extension_list: Option<ExtList>,
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
pub(super) struct SnapshotResource {
    pub(super) relationship_id: String,
    pub(super) target: SnapshotTarget,
}

/// Internal image bytes or an external linked image target.
///
/// External targets are retained as strings and are never fetched.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum SnapshotTarget {
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
    pub(super) part_name: &'a PackURI,
    pub(super) content_type: &'a str,
    pub(super) data: &'a Arc<Vec<u8>>,
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
    pub(super) id: String,
    pub(super) frozen: bool,
    pub(super) reference: Reference,
    pub(super) alternate_references: Vec<Reference>,
    pub(super) properties: Vec<Property>,
    pub(super) bindings: Vec<Binding>,
    pub(super) snapshot: Option<Snapshot>,
    pub(super) extension_list: Option<ExtList>,
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
    pub(super) dock_state: Dock,
    pub(super) visible: bool,
    pub(super) width: f64,
    pub(super) row: u32,
    pub(super) locked: bool,
    pub(super) relationship_id: String,
    pub(super) add_in: AddIn,
    pub(super) snapshot_resources: Vec<SnapshotResource>,
    pub(super) extension_list: Option<ExtList>,
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

    pub(super) fn next_snapshot_relationship_id(&self, base: &str) -> Result<String> {
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
    pub(super) panes: Vec<Pane>,
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

    pub(super) fn next_relationship_id(&self) -> Result<String> {
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

pub(super) fn canonicalize_pane_snapshot_resources(
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

pub(super) fn reconcile_snapshot_resource(
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
pub(super) struct OperationBudget {
    pub(super) xml_bytes: usize,
    pub(super) string_bytes: usize,
}

impl OperationBudget {
    pub(super) fn charge_xml(&mut self, bytes: usize, limits: &Limits) -> Result<()> {
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

    pub(super) fn charge_strings(&mut self, bytes: usize, limits: &Limits) -> Result<()> {
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

    pub(super) fn charge_authored(&mut self, xml: &[u8], limits: &Limits) -> Result<()> {
        self.charge_xml(xml.len(), limits)?;
        self.charge_strings(xml.len(), limits)
    }

    pub(super) fn charge_metadata(
        &mut self,
        bytes: usize,
        copies: usize,
        limits: &Limits,
    ) -> Result<()> {
        let retained = bytes.checked_mul(copies).ok_or(Error::Limit {
            resource: "indexed web extension package metadata bytes",
            max: limits.total_string_bytes,
            actual: usize::MAX,
        })?;
        self.charge_strings(retained, limits)
    }
}
