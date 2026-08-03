//! Typed WordprocessingML building blocks (the glossary / AutoText catalog).

use crate::{Error, Result};
use caseless::Caseless;
use litchi_opc::constants::content_type as ct;
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{ContentType, OpcPackage, PackURI};
use quick_xml::{
    Reader, XmlVersion,
    encoding::Decoder,
    events::{BytesStart, Event},
};
use std::collections::{HashMap, HashSet, VecDeque};
use std::sync::Arc;
use unicode_normalization::UnicodeNormalization;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WS: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const VML: &str = "urn:schemas-microsoft-com:vml";
const REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/glossaryDocument";
const STYLES_EFFECTS_REL: &str =
    "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects";
const CUSTOMIZATIONS_REL: &str =
    "http://schemas.microsoft.com/office/2006/relationships/keyMapCustomizations";
const CUSTOMIZATIONS_CT: &str = "application/vnd.ms-word.keyMapCustomizations+xml";
const ATTACHED_TOOLBARS_REL: &str =
    "http://schemas.microsoft.com/office/2006/relationships/attachedToolbars";
const ATTACHED_TOOLBARS_CT: &str = "application/vnd.ms-word.attachedToolbars";
const DIAGRAM_DRAWING_REL: &str =
    "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
const CHART_STYLE_REL: &str = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
const CHART_COLOR_STYLE_REL: &str =
    "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
const CHART_STYLE_REL_2012: &str =
    "http://schemas.microsoft.com/office/2012/relationships/chartStyle";
const CHART_COLOR_STYLE_REL_2012: &str =
    "http://schemas.microsoft.com/office/2012/relationships/chartColorStyle";
const ACTIVE_X_BINARY_REL: &str =
    "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary";
const STYLES_EFFECTS_CT: &str = "application/vnd.ms-word.stylesWithEffects+xml";
const CHART_STYLE_CT: &str = "application/vnd.ms-office.chartstyle+xml";
const CHART_COLOR_STYLE_CT: &str = "application/vnd.ms-office.chartcolorstyle+xml";
const ACTIVE_X_DESCRIPTOR_CT: &str = "application/vnd.ms-office.activeX+xml";
const ACTIVE_X_BINARY_CT: &str = "application/vnd.ms-office.activeX";
const RECIPIENT_DATA_CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.mailMergeRecipientData+xml";
const OBFUSCATED_FONT_CT: &str = "application/vnd.openxmlformats-officedocument.obfuscatedFont";
const FONT_DATA_CT: &str = "application/x-fontdata";
const FONT_TTF_CT: &str = "application/x-font-ttf";
const PRINTER_SETTINGS_CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.printerSettings";
const CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml";
const MAX: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_NODES: usize = 262_144;
const MAX_DOM_ATTRIBUTES: usize = 262_144;
const MAX_DOM_CONTENT: usize = 262_144;
const MAX_DOM_TOKENS: usize = 1_000_000;
const MAX_DOM_BYTES: usize = 64 * 1024 * 1024;
const MAX_PARTS: usize = 100_000;
const MAX_VALUES: usize = 4096;
const MAX_STRING: usize = 1024 * 1024;
const MAX_NAME_KEY: usize = 4 * MAX_STRING;
/// Conservative aggregate ceiling for inert glossary-owned auxiliary payloads.
const MAX_GRAPH_BYTES: usize = 256 * 1024 * 1024;
/// Aggregate ceiling for raw graph names, content types, and relationship metadata.
const MAX_GRAPH_METADATA_BYTES: usize = 32 * 1024 * 1024;

/// Namespace and relationship family used by a glossary part.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    fn index(self) -> usize {
        match self {
            Self::Transitional => 0,
            Self::Strict => 1,
        }
    }

    fn word(self) -> &'static str {
        match self {
            Self::Transitional => W,
            Self::Strict => WS,
        }
    }

    fn relationships(self) -> &'static str {
        match self {
            Self::Transitional => R,
            Self::Strict => RS,
        }
    }

    fn glossary_relationship(self) -> &'static str {
        match self {
            Self::Transitional => REL,
            Self::Strict => STRICT_REL,
        }
    }

    fn from_word(namespace: &str) -> Result<Self> {
        match namespace {
            W => Ok(Self::Transitional),
            WS => Ok(Self::Strict),
            _ => Err(invalid(
                "glossary root has an invalid WordprocessingML namespace",
            )),
        }
    }
}

bitflags::bitflags! {
    /// Compact set of building-block kinds.
    #[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
    pub struct Kind: u8 {
        const NONE = 1 << 0;
        const NORMAL = 1 << 1;
        const AUTO_EXPAND = 1 << 2;
        const TOOLBAR = 1 << 3;
        const SPELLER = 1 << 4;
        const FORM_FIELD = 1 << 5;
        const SDT_PLACEHOLDER = 1 << 6;
    }

    /// Compact set of insertion behaviors.
    #[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
    pub struct Insert: u8 {
        const CONTENT = 1 << 0;
        const PARAGRAPH = 1 << 1;
        const PAGE = 1 << 2;
    }
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Gallery(String);
impl Gallery {
    /// Validate a schema-defined building-block gallery token.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if GALLERIES.contains(&value.as_str()) {
            Ok(Self(value))
        } else {
            Err(invalid(format!("invalid document-part gallery '{value}'")))
        }
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Name {
    value: String,
    key: Arc<str>,
    decorated: Option<bool>,
}
impl Name {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_name(&value)?;
        let key = Arc::from(name_key(&value)?);
        Ok(Self {
            value,
            key,
            decorated: None,
        })
    }

    pub fn as_str(&self) -> &str {
        &self.value
    }

    pub fn decorated(&self) -> Option<bool> {
        self.decorated
    }

    fn key(&self) -> &str {
        &self.key
    }

    pub fn with_decorated(mut self, decorated: bool) -> Self {
        self.decorated = Some(decorated);
        self
    }
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Category {
    name: String,
    gallery: Gallery,
}
impl Category {
    pub fn new(name: impl Into<String>, gallery: Gallery) -> Result<Self> {
        let name = name.into();
        bounded(&name)?;
        if name.trim().is_empty() {
            return Err(invalid("building-block category name cannot be empty"));
        }
        Ok(Self { name, gallery })
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn gallery(&self) -> &Gallery {
        &self.gallery
    }
}

/// Canonical braced uppercase building-block GUID.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct Id(String);

impl Id {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if canonical_id(&value) {
            Ok(Self(value))
        } else {
            Err(invalid(format!(
                "building-block ID must be a braced uppercase GUID, got '{value}'"
            )))
        }
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Props {
    /// Optional in producer files; fresh authored entries require one.
    pub name: Option<Name>,
    pub style: Option<String>,
    pub category: Option<Category>,
    pub all_kinds: Option<bool>,
    pub kinds: Kind,
    pub inserts: Insert,
    pub description: Option<String>,
    pub id: Option<Id>,
}
impl Props {
    pub fn new(name: Name) -> Self {
        Self {
            name: Some(name),
            style: None,
            category: None,
            all_kinds: None,
            kinds: Kind::empty(),
            inserts: Insert::empty(),
            description: None,
            id: None,
        }
    }

    pub fn name(&self) -> Option<&Name> {
        self.name.as_ref()
    }

    pub fn with_name(mut self, name: Name) -> Self {
        self.name = Some(name);
        self
    }
}
#[derive(Clone, Debug)]
pub struct Entry {
    props: Option<Props>,
    /// Full inert `w:docPartBody` subtree.
    body: Option<Vec<u8>>,
    producer: Option<Box<ProducerEntry>>,
    sizes: [usize; 2],
    refs: Arc<[String]>,
    lineage: Option<Arc<Lineage>>,
}

#[derive(Clone, Debug)]
struct ProducerEntry {
    conformance: Conformance,
    xml: Arc<str>,
    refs: Arc<[String]>,
}

#[derive(Debug)]
struct Lineage;

struct EntryAnalysis {
    sizes: [usize; 2],
    refs: Arc<[String]>,
}
#[derive(Clone, Debug)]
pub struct Catalog {
    /// Full inert `w:background` subtree.
    background: Option<Vec<u8>>,
    background_refs: Arc<[String]>,
    background_lineage: Option<Arc<Lineage>>,
    entries: Vec<Entry>,
    /// Physical resource binding captured by package-aware loads.
    binding: Option<Box<Binding>>,
    state: CatalogState,
}

#[derive(Clone, Debug, Default)]
struct CatalogState {
    names: HashMap<Arc<str>, NameMatch>,
    entry_bytes: [usize; 2],
    background_bytes: [usize; 2],
}

#[derive(Clone, Copy, Debug)]
struct NameMatch {
    first: usize,
    count: usize,
}

/// Low-level physical glossary graph vocabulary.
///
/// Ordinary CRUD should use [`Catalog`] and [`load`]/[`put`]/[`remove`]. Raw
/// relationship IDs and part names are intentionally isolated here.
pub mod raw {
    use super::{Catalog, Conformance};
    use std::sync::Arc;

    #[derive(Clone, Debug, PartialEq, Eq)]
    pub struct Rel {
        pub id: String,
        pub kind: String,
        pub target: String,
        pub external: bool,
    }

    #[derive(Clone, Debug, PartialEq, Eq)]
    pub struct Part {
        pub name: String,
        pub content_type: String,
        pub(super) data: Arc<Vec<u8>>,
        pub rels: Vec<Rel>,
    }

    impl Part {
        /// Validate basic physical metadata and take ownership of a payload.
        pub fn new(
            name: impl Into<String>,
            content_type: impl Into<String>,
            data: Vec<u8>,
        ) -> super::Result<Self> {
            let name = name.into();
            let content_type = content_type.into();
            super::validate_raw_part(&name, &content_type, data.len())?;
            Ok(Self {
                name,
                content_type,
                data: Arc::new(data),
                rels: Vec::new(),
            })
        }

        pub fn data(&self) -> &[u8] {
            self.data.as_slice()
        }

        /// Consume the part and retain shared ownership of its payload.
        pub fn into_data(self) -> Arc<Vec<u8>> {
            self.data
        }

        pub(super) fn shared_data(&self) -> Arc<Vec<u8>> {
            Arc::clone(&self.data)
        }

        pub(super) fn from_shared(
            name: String,
            content_type: String,
            data: Arc<Vec<u8>>,
            rels: Vec<Rel>,
        ) -> Self {
            Self {
                name,
                content_type,
                data,
                rels,
            }
        }
    }

    #[derive(Clone, Debug, PartialEq, Eq)]
    pub struct Graph {
        pub catalog: Catalog,
        pub conformance: Conformance,
        pub rels: Vec<Rel>,
        pub parts: Vec<Part>,
        pub(super) root_name: String,
        pub(super) root_xml: Option<Arc<Vec<u8>>>,
        pub(super) owner_main: Option<String>,
        pub(super) owner_id: Option<String>,
        pub(super) owner_target: Option<String>,
    }

    impl Graph {
        pub fn new(catalog: Catalog, conformance: Conformance) -> Self {
            Self {
                catalog,
                conformance,
                rels: Vec::new(),
                parts: Vec::new(),
                root_name: "/word/glossary/document.xml".to_owned(),
                root_xml: None,
                owner_main: None,
                owner_id: None,
                owner_target: None,
            }
        }

        /// Producer-selected glossary root part name.
        pub fn root_name(&self) -> &str {
            &self.root_name
        }

        /// Original producer XML, when this graph was loaded from a package.
        pub fn root_xml(&self) -> Option<&[u8]> {
            self.root_xml.as_deref().map(Vec::as_slice)
        }

        /// Producer-selected ID of the main-document owner relationship.
        pub fn owner_relationship_id(&self) -> Option<&str> {
            self.owner_id.as_deref()
        }

        /// Producer-selected main-document part that owns this graph.
        pub fn owner_main_part(&self) -> Option<&str> {
            self.owner_main.as_deref()
        }

        /// Producer-selected literal target of the main-document owner relationship.
        pub fn owner_target(&self) -> Option<&str> {
            self.owner_target.as_deref()
        }
    }

    /// Load the complete glossary-owned OPC graph.
    pub fn load(package: &litchi_opc::OpcPackage) -> super::Result<Option<Graph>> {
        super::load_graph(package)
    }

    /// Publish a complete graph without consuming the caller's recovery copy.
    pub fn put(package: &mut litchi_opc::OpcPackage, graph: &Graph) -> super::Result<bool> {
        super::put_graph(package, graph)
    }

    /// Remove and return the complete graph for graph-preserving transfer.
    pub fn remove(package: &mut litchi_opc::OpcPackage) -> super::Result<Option<Graph>> {
        super::remove_graph(package)
    }
}

#[derive(Clone, Debug)]
struct Binding {
    conformance: Conformance,
    rels: Vec<raw::Rel>,
    parts: Vec<raw::Part>,
    root_name: String,
    owner_main: Option<String>,
    owner_id: Option<String>,
    owner_target: Option<String>,
    lineage: Arc<Lineage>,
}

impl Binding {
    fn from_graph(graph: &raw::Graph) -> Self {
        Self {
            conformance: graph.conformance,
            rels: graph.rels.clone(),
            parts: graph.parts.clone(),
            root_name: graph.root_name.clone(),
            owner_main: graph.owner_main.clone(),
            owner_id: graph.owner_id.clone(),
            owner_target: graph.owner_target.clone(),
            lineage: Arc::new(Lineage),
        }
    }

    fn matches(&self, graph: &raw::Graph) -> bool {
        self.conformance == graph.conformance
            && self.rels == graph.rels
            && self.parts.len() == graph.parts.len()
            && self.parts.iter().zip(&graph.parts).all(|(left, right)| {
                left.name == right.name
                    && left.content_type == right.content_type
                    && left.rels == right.rels
                    && (Arc::ptr_eq(&left.data, &right.data) || left.data == right.data)
            })
            && self.root_name == graph.root_name
            && self.owner_main == graph.owner_main
            && self.owner_id == graph.owner_id
            && self.owner_target == graph.owner_target
    }
}

impl PartialEq for Catalog {
    fn eq(&self, other: &Self) -> bool {
        self.background == other.background && self.entries == other.entries
    }
}

impl Eq for Catalog {}

impl PartialEq for Entry {
    fn eq(&self, other: &Self) -> bool {
        self.props == other.props && self.body == other.body
    }
}

impl Eq for Entry {}

impl CatalogState {
    fn offset(&self, name: &str) -> Result<Option<usize>> {
        let key = name_key(name)?;
        self.offset_key(&key, name)
    }

    fn offset_key(&self, key: &str, display: &str) -> Result<Option<usize>> {
        match self.names.get(key) {
            None => Ok(None),
            Some(found) if found.count == 1 => Ok(Some(found.first)),
            Some(_) => Err(Error::Ambiguous {
                object: "building block",
                key: display.to_owned(),
            }),
        }
    }

    fn rebuild_names(&mut self, entries: &[Entry]) -> Result<()> {
        let mut names = HashMap::new();
        names
            .try_reserve(entries.len())
            .map_err(|source| Error::Allocation {
                resource: "glossary semantic-name index",
                source,
            })?;
        for (index, entry) in entries.iter().enumerate() {
            let Some(name) = entry.props.as_ref().and_then(|props| props.name.as_ref()) else {
                continue;
            };
            if let Some(found) = names.get_mut(name.key()) {
                let found: &mut NameMatch = found;
                found.count = found
                    .count
                    .checked_add(1)
                    .ok_or_else(|| invalid("glossary duplicate-name count overflow"))?;
            } else {
                names.insert(
                    Arc::clone(&name.key),
                    NameMatch {
                        first: index,
                        count: 1,
                    },
                );
            }
        }
        self.names = names;
        Ok(())
    }

    fn validate_entry_key(&self, key: Option<&str>) -> Result<()> {
        if key.is_some_and(|key| !self.names.contains_key(key)) {
            Err(invalid("semantic-name index does not cover the catalog"))
        } else {
            Ok(())
        }
    }

    fn validate_replacement_key(
        &self,
        index: usize,
        old_key: Option<&str>,
        new_key: &str,
        display: &str,
    ) -> Result<()> {
        self.validate_entry_key(old_key)?;
        if old_key != Some(new_key) && self.names.contains_key(new_key) {
            return Err(invalid(format!(
                "building block '{display}' already exists"
            )));
        }
        if old_key == Some(new_key)
            && self
                .names
                .get(new_key)
                .is_some_and(|found| found.count == 1 && found.first != index)
        {
            return Err(invalid("semantic-name index points to a different entry"));
        }
        Ok(())
    }

    fn replace_name(
        &mut self,
        index: usize,
        old_key: Option<&str>,
        new_key: Arc<str>,
        entries: &[Entry],
    ) {
        if old_key == Some(new_key.as_ref()) {
            return;
        }
        if let Some(old_key) = old_key {
            let old = self.names.get(old_key).copied();
            if let Some(old) = old {
                if old.count == 1 {
                    self.names.remove(old_key);
                } else if let Some(found) = self.names.get_mut(old_key) {
                    found.count -= 1;
                    if found.first == index {
                        found.first = entries
                            .iter()
                            .position(|entry| entry_key(entry) == Some(old_key))
                            .unwrap_or(found.first);
                    }
                }
            }
        }
        self.names.insert(
            new_key,
            NameMatch {
                first: index,
                count: 1,
            },
        );
    }

    fn remove_name(&mut self, index: usize, old_key: Option<&str>, entries: &[Entry]) {
        let old = old_key.and_then(|key| self.names.get(key).copied());
        let next = old_key.and_then(|key| {
            entries
                .iter()
                .position(|entry| entry_key(entry) == Some(key))
        });
        for found in self.names.values_mut() {
            if found.first > index {
                found.first -= 1;
            }
        }
        if let (Some(old_key), Some(old)) = (old_key, old) {
            if old.count == 1 {
                self.names.remove(old_key);
            } else if let Some(found) = self.names.get_mut(old_key) {
                found.count -= 1;
                if old.first == index {
                    found.first = next.unwrap_or(found.first);
                }
            }
        }
    }

    fn refresh_name_positions(&mut self, entries: &[Entry]) -> Result<()> {
        if entries.len() > MAX_PARTS
            || entries.iter().any(|entry| {
                entry
                    .props
                    .as_ref()
                    .and_then(|props| props.name.as_ref())
                    .is_some_and(|name| !self.names.contains_key(name.key()))
            })
        {
            return Err(invalid("semantic-name index does not cover the catalog"));
        }
        for found in self.names.values_mut() {
            found.first = usize::MAX;
            found.count = 0;
        }
        for (index, entry) in entries.iter().enumerate() {
            let Some(name) = entry.props.as_ref().and_then(|props| props.name.as_ref()) else {
                continue;
            };
            if let Some(found) = self.names.get_mut(name.key()) {
                if found.count == 0 {
                    found.first = index;
                }
                found.count = found.count.saturating_add(1);
            }
        }
        self.names.retain(|_, found| found.count != 0);
        Ok(())
    }
}

impl Catalog {
    fn rebuild_state(&mut self) -> Result<()> {
        let mut state = CatalogState::default();
        state.rebuild_names(&self.entries)?;
        let background = background_analysis(self.background.as_deref())?;
        state.background_bytes = background.sizes;
        self.background_refs = background.refs;
        for entry in &mut self.entries {
            let analysis = entry_analysis(entry)?;
            entry.sizes = analysis.sizes;
            entry.refs = analysis.refs;
            state.entry_bytes = add_sizes(state.entry_bytes, entry.sizes)?;
        }
        validate_catalog_sizes(
            state.entry_bytes,
            state.background_bytes,
            self.entries.len(),
        )?;
        self.state = state;
        Ok(())
    }

    pub fn new() -> Self {
        Self {
            background: None,
            background_refs: Arc::from([]),
            background_lineage: None,
            entries: Vec::new(),
            binding: None,
            state: CatalogState::default(),
        }
    }

    pub fn background(&self) -> Option<&[u8]> {
        self.background.as_deref()
    }

    fn validate_entry_lineage(&self, entry: &Entry) -> Result<()> {
        if !entry.has_relationship_references() || self.binding.is_none() {
            return Ok(());
        }
        let matches = self
            .binding
            .as_deref()
            .zip(entry.lineage.as_ref())
            .is_some_and(|(binding, lineage)| Arc::ptr_eq(&binding.lineage, lineage));
        if matches {
            Ok(())
        } else {
            Err(invalid(
                "relationship-bearing building block belongs to another physical graph; use glossary::raw for graph transfer",
            ))
        }
    }

    fn validate_bound_lineages(&self) -> Result<()> {
        for entry in &self.entries {
            self.validate_entry_lineage(entry)?;
        }
        if !self.background_refs.is_empty() && self.binding.is_some() {
            let matches = self
                .binding
                .as_deref()
                .zip(self.background_lineage.as_ref())
                .is_some_and(|(binding, lineage)| Arc::ptr_eq(&binding.lineage, lineage));
            if !matches {
                return Err(invalid(
                    "relationship-bearing background belongs to another physical graph; use glossary::raw for graph transfer",
                ));
            }
        }
        Ok(())
    }

    pub fn set_background(&mut self, xml: Vec<u8>) -> Result<Option<Vec<u8>>> {
        if self.background.as_ref() == Some(&xml) {
            return Ok(self.background.replace(xml));
        }
        let analysis = background_analysis(Some(&xml))?;
        if self.binding.is_some() && !analysis.refs.is_empty() {
            return Err(invalid(
                "relationship-bearing background requires glossary::raw graph publication",
            ));
        }
        validate_catalog_sizes(self.state.entry_bytes, analysis.sizes, self.entries.len())?;
        let previous = self.background.replace(xml);
        self.background_refs = analysis.refs;
        self.background_lineage = None;
        self.state.background_bytes = analysis.sizes;
        Ok(previous)
    }

    pub fn clear_background(&mut self) -> Option<Vec<u8>> {
        self.state.background_bytes = [0; 2];
        self.background_refs = Arc::from([]);
        self.background_lineage = None;
        self.background.take()
    }

    pub fn entries(&self) -> &[Entry] {
        &self.entries
    }

    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Entry> {
        self.entries.iter()
    }

    pub fn len(&self) -> usize {
        self.entries.len()
    }

    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Find a uniquely named entry using Unicode default case folding.
    pub fn get(&self, name: &str) -> Result<Option<&Entry>> {
        Ok(self
            .state
            .offset(name)?
            .and_then(|offset| self.entries.get(offset)))
    }

    /// Checked numeric fallback for import and inspection workflows.
    pub fn at(&self, index: usize) -> Result<&Entry> {
        self.entries.get(index).ok_or(Error::OutOfBounds {
            object: "glossary entry",
            index,
            len: self.entries.len(),
        })
    }

    /// Add a fresh, uniquely named entry by moving it into the catalog.
    pub fn add(&mut self, entry: Entry) -> Result<usize> {
        validate_authored_name(&entry)?;
        self.validate_entry_lineage(&entry)?;
        let name = entry
            .name()
            .ok_or_else(|| invalid("authored building block requires a name"))?;
        let key = Arc::clone(
            &entry
                .props
                .as_ref()
                .and_then(|props| props.name.as_ref())
                .ok_or_else(|| invalid("authored building block requires a name"))?
                .key,
        );
        if self.state.offset_key(key.as_ref(), name)?.is_some() {
            return Err(invalid(format!("building block '{name}' already exists")));
        }
        if self.entries.len() >= MAX_PARTS {
            return Err(invalid("glossary entry limit exceeded"));
        }
        self.entries
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "glossary entry insertion",
                source,
            })?;
        self.state
            .names
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "glossary semantic-name index",
                source,
            })?;
        let candidate_bytes = add_sizes(self.state.entry_bytes, entry.sizes)?;
        validate_catalog_sizes(
            candidate_bytes,
            self.state.background_bytes,
            self.entries.len() + 1,
        )?;
        let index = self.entries.len();
        self.entries.push(entry);
        self.state.names.insert(
            key,
            NameMatch {
                first: index,
                count: 1,
            },
        );
        self.state.entry_bytes = candidate_bytes;
        Ok(index)
    }

    /// Insert or replace an entry selected by the replacement's semantic name.
    pub fn put(&mut self, entry: Entry) -> Result<Option<Entry>> {
        let name = entry
            .name()
            .ok_or_else(|| invalid("authored building block requires a name"))?;
        let key = entry
            .props
            .as_ref()
            .and_then(|props| props.name.as_ref())
            .map(Name::key)
            .ok_or_else(|| invalid("authored building block requires a name"))?;
        let offset = self.state.offset_key(key, name)?;
        if let Some(offset) = offset {
            return self.replace_at(offset, entry).map(Some);
        }
        self.add(entry)?;
        Ok(None)
    }

    /// Replace a uniquely named entry while selecting independently of its new name.
    pub fn replace(&mut self, name: &str, entry: Entry) -> Result<Option<Entry>> {
        let Some(index) = self.state.offset(name)? else {
            return Ok(None);
        };
        self.replace_at(index, entry).map(Some)
    }

    /// Checked numeric fallback for replacement and ambiguous-name repair.
    pub fn replace_at(&mut self, index: usize, entry: Entry) -> Result<Entry> {
        validate_authored_name(&entry)?;
        self.validate_entry_lineage(&entry)?;
        let previous = self.entries.get(index).ok_or(Error::OutOfBounds {
            object: "glossary entry",
            index,
            len: self.entries.len(),
        })?;
        let new_name = entry
            .name()
            .ok_or_else(|| invalid("authored building block requires a name"))?;
        let new_key = Arc::clone(
            &entry
                .props
                .as_ref()
                .and_then(|props| props.name.as_ref())
                .ok_or_else(|| invalid("authored building block requires a name"))?
                .key,
        );
        self.state.validate_replacement_key(
            index,
            entry_key(previous),
            new_key.as_ref(),
            new_name,
        )?;
        self.state
            .names
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "glossary semantic-name index",
                source,
            })?;
        let candidate_bytes = replace_sizes(self.state.entry_bytes, previous.sizes, entry.sizes)?;
        validate_catalog_sizes(
            candidate_bytes,
            self.state.background_bytes,
            self.entries.len(),
        )?;
        let slot = self
            .entries
            .get_mut(index)
            .ok_or_else(|| invalid("glossary replacement index changed"))?;
        let previous = std::mem::replace(slot, entry);
        self.state
            .replace_name(index, entry_key(&previous), new_key, &self.entries);
        self.state.entry_bytes = candidate_bytes;
        Ok(previous)
    }

    /// Rename a uniquely named entry without copying its body.
    pub fn rename(&mut self, name: &str, replacement: Name) -> Result<bool> {
        let Some(index) = self.state.offset(name)? else {
            return Ok(false);
        };
        self.rename_at(index, replacement)
    }

    /// Checked numeric fallback for renaming an ambiguous producer entry.
    pub fn rename_at(&mut self, index: usize, replacement: Name) -> Result<bool> {
        let entry = self.entries.get(index).ok_or(Error::OutOfBounds {
            object: "glossary entry",
            index,
            len: self.entries.len(),
        })?;
        let current = entry
            .props
            .as_ref()
            .and_then(|props| props.name.as_ref())
            .ok_or_else(|| invalid("building block has no name to rename"))?;
        if current == &replacement {
            return Ok(false);
        }
        let old_key = Arc::clone(&current.key);
        let new_key = Arc::clone(&replacement.key);
        self.state.validate_replacement_key(
            index,
            Some(old_key.as_ref()),
            new_key.as_ref(),
            replacement.as_str(),
        )?;
        self.state
            .names
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "glossary semantic-name index",
                source,
            })?;
        let old_sizes = entry.sizes;
        let (previous_name, producer) = {
            let entry = self
                .entries
                .get_mut(index)
                .ok_or_else(|| invalid("glossary rename index changed"))?;
            let props = entry
                .props
                .as_mut()
                .ok_or_else(|| invalid("building block has no properties to rename"))?;
            let previous_name = props
                .name
                .replace(replacement)
                .ok_or_else(|| invalid("building block name disappeared during rename"))?;
            let producer = entry.producer.take();
            (previous_name, producer)
        };
        let analysis = match entry_analysis(
            self.entries
                .get(index)
                .ok_or_else(|| invalid("glossary rename index changed"))?,
        ) {
            Ok(analysis) => analysis,
            Err(error) => {
                restore_name(
                    self.entries
                        .get_mut(index)
                        .ok_or_else(|| invalid("glossary rename index changed"))?,
                    previous_name,
                    producer,
                )?;
                return Err(error);
            },
        };
        let candidate_bytes = match replace_sizes(self.state.entry_bytes, old_sizes, analysis.sizes)
            .and_then(|bytes| {
                validate_catalog_sizes(bytes, self.state.background_bytes, self.entries.len())?;
                Ok(bytes)
            }) {
            Ok(bytes) => bytes,
            Err(error) => {
                restore_name(
                    self.entries
                        .get_mut(index)
                        .ok_or_else(|| invalid("glossary rename index changed"))?,
                    previous_name,
                    producer,
                )?;
                return Err(error);
            },
        };
        if let Some(entry) = self.entries.get_mut(index) {
            entry.sizes = analysis.sizes;
            entry.refs = analysis.refs;
        }
        self.state
            .replace_name(index, Some(old_key.as_ref()), new_key, &self.entries);
        self.state.entry_bytes = candidate_bytes;
        Ok(true)
    }

    pub fn remove(&mut self, name: &str) -> Result<Option<Entry>> {
        let Some(offset) = self.state.offset(name)? else {
            return Ok(None);
        };
        self.remove_at(offset).map(Some)
    }

    /// Checked numeric fallback for removal.
    pub fn remove_at(&mut self, index: usize) -> Result<Entry> {
        if index >= self.entries.len() {
            return Err(Error::OutOfBounds {
                object: "glossary entry",
                index,
                len: self.entries.len(),
            });
        }
        let entry = self
            .entries
            .get(index)
            .ok_or_else(|| invalid("missing entry"))?;
        self.state.validate_entry_key(entry_key(entry))?;
        let sizes = entry.sizes;
        let candidate_bytes = [
            self.state.entry_bytes[0]
                .checked_sub(sizes[0])
                .ok_or_else(|| invalid("Transitional glossary size underflow"))?,
            self.state.entry_bytes[1]
                .checked_sub(sizes[1])
                .ok_or_else(|| invalid("Strict glossary size underflow"))?,
        ];
        let removed = self.entries.remove(index);
        self.state
            .remove_name(index, entry_key(&removed), &self.entries);
        self.state.entry_bytes = candidate_bytes;
        Ok(removed)
    }

    pub fn move_to(&mut self, name: &str, to: usize) -> Result<bool> {
        let Some(from) = self.state.offset(name)? else {
            return Ok(false);
        };
        self.move_at(from, to)?;
        Ok(true)
    }

    /// Checked numeric fallback for reordering.
    pub fn move_at(&mut self, from: usize, to: usize) -> Result<bool> {
        if from >= self.entries.len() || to >= self.entries.len() {
            let index = if from >= self.entries.len() { from } else { to };
            return Err(Error::OutOfBounds {
                object: "glossary entry",
                index,
                len: self.entries.len(),
            });
        }
        if from != to {
            let entry = self.entries.remove(from);
            self.entries.insert(to, entry);
            if let Err(error) = self.state.refresh_name_positions(&self.entries) {
                let entry = self.entries.remove(to);
                self.entries.insert(from, entry);
                return Err(error);
            }
        }
        Ok(from != to)
    }

    pub fn clear(&mut self) -> usize {
        let count = self.entries.len();
        self.entries.clear();
        self.state.names.clear();
        self.state.entry_bytes = [0; 2];
        count
    }
}

impl Default for Catalog {
    fn default() -> Self {
        Self::new()
    }
}

fn entry_key(entry: &Entry) -> Option<&str> {
    entry
        .props
        .as_ref()
        .and_then(|props| props.name.as_ref())
        .map(Name::key)
}

fn restore_name(entry: &mut Entry, name: Name, producer: Option<Box<ProducerEntry>>) -> Result<()> {
    let props = entry
        .props
        .as_mut()
        .ok_or_else(|| invalid("building block properties disappeared during rename rollback"))?;
    props.name = Some(name);
    entry.producer = producer;
    Ok(())
}

impl Entry {
    fn has_relationship_references(&self) -> bool {
        !self.refs.is_empty()
            || self
                .producer
                .as_deref()
                .is_some_and(|producer| !producer.refs.is_empty())
    }

    pub fn new(name: impl Into<String>, body_xml: Vec<u8>) -> Result<Self> {
        let name = Name::new(name)?;
        let entry = Self {
            props: Some(Props::new(name)),
            body: Some(body_xml),
            producer: None,
            sizes: [0; 2],
            refs: Arc::from([]),
            lineage: None,
        };
        let mut entry = entry;
        let analysis = authored_analysis(&entry)?;
        entry.sizes = analysis.sizes;
        entry.refs = analysis.refs;
        Ok(entry)
    }

    pub fn name(&self) -> Option<&str> {
        self.props
            .as_ref()
            .and_then(|props| props.name.as_ref())
            .map(Name::as_str)
    }

    pub fn props(&self) -> Option<&Props> {
        self.props.as_ref()
    }

    pub fn body(&self) -> Option<&[u8]> {
        self.body.as_deref()
    }

    pub fn with_props(mut self, props: Props) -> Result<Self> {
        if self.props.as_ref() != Some(&props) {
            self.producer = None;
        }
        self.props = Some(props);
        let analysis = authored_analysis(&self)?;
        self.sizes = analysis.sizes;
        self.refs = analysis.refs;
        Ok(self)
    }

    pub fn with_body(mut self, body: Vec<u8>) -> Result<Self> {
        if self.body.as_ref() != Some(&body) {
            self.producer = None;
            self.lineage = None;
        }
        self.body = Some(body);
        let analysis = authored_analysis(&self)?;
        self.sizes = analysis.sizes;
        self.refs = analysis.refs;
        Ok(self)
    }

    pub fn into_body(self) -> Option<Vec<u8>> {
        self.body
    }
}

/// Parse bounded glossary XML and report its namespace dialect.
pub fn read(xml: &[u8]) -> Result<(Catalog, Conformance)> {
    if xml.len() > MAX {
        return Err(invalid("glossary document exceeds 32 MiB"));
    }
    let original = parse_dom(xml)?;
    let original_conformance = Conformance::from_word(original.ns.as_ref())?;
    let producer_entries = extract_producer_entries(original, original_conformance)?;
    let limits = litchi_ooxml_common::MceLimits {
        max_input_bytes: MAX,
        max_output_bytes: MAX,
        max_depth: MAX_DEPTH,
        max_namespace_bindings: MAX_VALUES,
        max_directive_tokens: MAX_VALUES,
        max_choices_per_alternate: MAX_VALUES,
    };
    let xml = litchi_ooxml_common::process_markup_compatibility(
        xml,
        &litchi_ooxml_common::MceCapabilities::default(),
        &limits,
    )?
    .xml;
    if xml.len() > MAX {
        return Err(invalid("processed glossary document exceeds 32 MiB"));
    }
    let root = parse_dom(xml.as_ref())?;
    let conformance = Conformance::from_word(root.ns.as_ref())?;
    if conformance != original_conformance {
        return Err(invalid(
            "MCE preprocessing changed the glossary conformance family",
        ));
    }
    validate_word_dialect(&root, conformance)?;
    let mut catalog = project(&root)?;
    drop(root);
    attach_producer_entries(&mut catalog, producer_entries);
    catalog.rebuild_state()?;
    Ok((catalog, conformance))
}

/// Serialize a catalog canonically in the selected dialect.
pub fn write(value: &Catalog, conformance: Conformance) -> Result<Vec<u8>> {
    let plan = plan_write(value, conformance)?;
    let mut xml = String::new();
    xml.try_reserve_exact(plan.bytes)
        .map_err(|source| Error::Allocation {
            resource: "glossary XML",
            source,
        })?;
    emit_catalog(&mut xml, value, conformance, &plan)?;
    if xml.len() != plan.bytes {
        return Err(invalid("glossary XML write plan did not match output"));
    }
    Ok(xml.into_bytes())
}

#[derive(Clone, Debug)]
struct Owner {
    main: PackURI,
    root: PackURI,
    relationship_id: String,
    relationship_target: String,
    conformance: Conformance,
}

/// Load the semantic catalog and its namespace dialect, without copying auxiliaries.
pub fn load(package: &OpcPackage) -> Result<Option<(Catalog, Conformance)>> {
    let Some(graph) = load_graph(package)? else {
        return Ok(None);
    };
    let binding = Binding::from_graph(&graph);
    let conformance = graph.conformance;
    let mut catalog = graph.catalog;
    for entry in &mut catalog.entries {
        if entry.has_relationship_references() {
            entry.lineage = Some(Arc::clone(&binding.lineage));
        }
    }
    if !catalog.background_refs.is_empty() {
        catalog.background_lineage = Some(Arc::clone(&binding.lineage));
    }
    catalog.binding = Some(Box::new(binding));
    Ok(Some((catalog, conformance)))
}

/// Move a semantic catalog into the package while preserving its auxiliary graph.
///
/// An unchanged package-loaded catalog is a byte- and signature-preserving no-op.
pub fn put(
    package: &mut OpcPackage,
    mut catalog: Catalog,
    conformance: Conformance,
) -> Result<bool> {
    validate_package_conformance(package, conformance)?;
    catalog.validate_bound_lineages()?;
    let existing = load_graph(package)?;
    let binding = catalog.binding.take();
    let bound_to_destination = binding
        .as_deref()
        .zip(existing.as_ref())
        .is_some_and(|(binding, graph)| binding.matches(graph));
    if bound_to_destination
        && existing
            .as_ref()
            .is_some_and(|graph| graph.catalog == catalog)
    {
        return Ok(false);
    }
    if !catalog_relationship_references(&catalog, conformance)?.is_empty() && !bound_to_destination
    {
        return Err(invalid(
            "glossary relationship references are bound to another physical graph; use glossary::raw for graph transfer",
        ));
    }
    let mut graph = if bound_to_destination {
        existing.ok_or_else(|| invalid("bound glossary graph disappeared"))?
    } else {
        let mut graph = raw::Graph::new(Catalog::new(), conformance);
        seed_semantic_graph(package, &mut graph)?;
        graph
    };
    graph.catalog = catalog;
    graph.conformance = conformance;
    put_graph(package, &graph)
}

fn seed_semantic_graph(package: &OpcPackage, graph: &mut raw::Graph) -> Result<()> {
    if !graph.rels.is_empty() || !graph.parts.is_empty() {
        return Err(invalid("semantic glossary seed graph is not empty"));
    }
    let root_uri = free_part_name(
        package,
        "/word/glossary/document.xml",
        "/word/glossary/document%d.xml",
    )?;
    graph.root_name = root_uri.as_str().to_owned();
    let namespace = graph.conformance.word();
    for (index, kind, preferred, template, content_type, root) in [
        (
            1,
            "styles",
            "/word/glossary/styles.xml",
            "/word/glossary/styles%d.xml",
            ct::WML_STYLES,
            "styles",
        ),
        (
            2,
            "settings",
            "/word/glossary/settings.xml",
            "/word/glossary/settings%d.xml",
            ct::WML_SETTINGS,
            "settings",
        ),
        (
            3,
            "fontTable",
            "/word/glossary/fontTable.xml",
            "/word/glossary/fontTable%d.xml",
            ct::WML_FONT_TABLE,
            "fonts",
        ),
        (
            4,
            "webSettings",
            "/word/glossary/webSettings.xml",
            "/word/glossary/webSettings%d.xml",
            ct::WML_WEB_SETTINGS,
            "webSettings",
        ),
    ] {
        let part_uri = free_part_name(package, preferred, template)?;
        graph.rels.push(raw::Rel {
            id: format!("rId{index}"),
            kind: format!("{}/{kind}", graph.conformance.relationships()),
            target: part_uri.relative_ref(root_uri.base_uri()),
            external: false,
        });
        let source = package
            .main_document_part()?
            .rels()
            .iter()
            .find_map(|relationship| {
                (!relationship.is_external()
                    && relationship_kind(graph.conformance, relationship.reltype()) == Some(kind))
                .then(|| relationship.target_partname().ok())
                .flatten()
                .and_then(|target| package.get_part(&target).ok())
                .filter(|part| {
                    part.content_type() == content_type && part.rels().iter().next().is_none()
                })
                .map(Part::blob_arc)
            });
        let data = source.unwrap_or_else(|| {
            Arc::new(format!(r#"<w:{root} xmlns:w="{namespace}"/>"#).into_bytes())
        });
        graph.parts.push(raw::Part::from_shared(
            part_uri.as_str().to_owned(),
            content_type.to_owned(),
            data,
            Vec::new(),
        ));
    }
    Ok(())
}

fn free_part_name(package: &OpcPackage, preferred: &str, template: &str) -> Result<PackURI> {
    let preferred = PackURI::new(preferred).map_err(Error::Uri)?;
    if package.validate_new_part_name(&preferred).is_ok() {
        return Ok(preferred);
    }
    let marker = template
        .find("%d")
        .ok_or_else(|| invalid("glossary part-name template is missing '%d'"))?;
    for index in 1..=10_000u32 {
        let mut candidate = String::new();
        candidate
            .try_reserve(template.len().saturating_add(10))
            .map_err(|source| Error::Allocation {
                resource: "glossary part name",
                source,
            })?;
        candidate.push_str(&template[..marker]);
        candidate.push_str(&index.to_string());
        candidate.push_str(&template[marker + 2..]);
        let candidate = PackURI::new(&candidate).map_err(Error::Uri)?;
        if package.validate_new_part_name(&candidate).is_ok() {
            return Ok(candidate);
        }
    }
    Err(invalid(
        "glossary part-name allocation exhausted 10,000 candidates",
    ))
}

/// Remove the glossary graph. Absence is a signature-preserving no-op.
///
/// Use [`raw::remove`] to move the complete physical graph elsewhere.
pub fn remove(package: &mut OpcPackage) -> Result<bool> {
    Ok(remove_graph(package)?.is_some())
}

fn load_graph(package: &OpcPackage) -> Result<Option<raw::Graph>> {
    let Some(owner) = locate(package)? else {
        return Ok(None);
    };
    let root = package.get_part(&owner.root)?;
    let (catalog, conformance) = read(root.blob())?;
    if conformance != owner.conformance {
        return Err(invalid(
            "glossary relationship and XML use different conformance families",
        ));
    }
    validate_relationship_integrity(&catalog, root, conformance)?;
    let owned = glossary_owned_parts(package, &owner.root, conformance)?;
    validate_exclusive_ownership(package, &owner, &owned)?;
    let rels = copy_relationships(root)?;
    let mut names = owned
        .into_iter()
        .filter(|uri| uri != &owner.root)
        .collect::<Vec<_>>();
    names.sort_by(|left, right| left.as_str().cmp(right.as_str()));
    let mut parts = Vec::new();
    parts
        .try_reserve_exact(names.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary auxiliary graph",
            source,
        })?;
    let mut total_bytes = 0usize;
    for uri in names {
        let part = package.get_part(&uri)?;
        total_bytes = total_bytes
            .checked_add(part.blob().len())
            .ok_or_else(|| invalid("glossary auxiliary payload size overflow"))?;
        if total_bytes > MAX_GRAPH_BYTES {
            return Err(invalid(
                "glossary auxiliary graph exceeds the 256 MiB aggregate limit",
            ));
        }
        parts.push(raw::Part::from_shared(
            uri.as_str().to_owned(),
            part.content_type().to_owned(),
            part.blob_arc(),
            copy_relationships(part)?,
        ));
    }
    Ok(Some(raw::Graph {
        catalog,
        conformance,
        rels,
        parts,
        root_name: owner.root.as_str().to_owned(),
        root_xml: Some(root.blob_arc()),
        owner_main: Some(owner.main.as_str().to_owned()),
        owner_id: Some(owner.relationship_id),
        owner_target: Some(owner.relationship_target),
    }))
}

fn copy_relationships(part: &dyn Part) -> Result<Vec<raw::Rel>> {
    let relationship_count = part.rels().iter().count();
    let mut rels = Vec::new();
    rels.try_reserve_exact(relationship_count)
        .map_err(|source| Error::Allocation {
            resource: "glossary relationships",
            source,
        })?;
    for relationship in part.rels().iter() {
        rels.push(raw::Rel {
            id: relationship.r_id().to_owned(),
            kind: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            external: relationship.is_external(),
        });
    }
    rels.sort_by(|left, right| left.id.cmp(&right.id));
    Ok(rels)
}

fn put_graph(package: &mut OpcPackage, value: &raw::Graph) -> Result<bool> {
    validate_package_conformance(package, value.conformance)?;
    validate_raw_graph_metadata(value)?;
    let preserve_root = if let Some(xml) = &value.root_xml {
        let (source, source_conformance) = read(xml.as_slice())?;
        source_conformance == value.conformance && source == value.catalog
    } else {
        false
    };
    let root_uri = PackURI::new(&value.root_name).map_err(Error::Uri)?;
    if is_signature_part(&root_uri) || is_reserved_physical_part(&root_uri) {
        return Err(invalid(
            "glossary root cannot use reserved OPC package infrastructure",
        ));
    }
    let mut root = if preserve_root {
        BlobPart::new_shared(
            root_uri.clone(),
            CT.to_owned(),
            value
                .root_xml
                .as_ref()
                .ok_or_else(|| invalid("preserved glossary XML is missing"))?
                .clone(),
        )
    } else {
        BlobPart::new(
            root_uri.clone(),
            CT.to_owned(),
            write(&value.catalog, value.conformance)?,
        )
    };
    let canonical_catalog;
    let effective_catalog = if preserve_root {
        &value.catalog
    } else {
        let (canonical, canonical_conformance) = read(root.blob())?;
        if canonical_conformance != value.conformance {
            return Err(invalid("canonical glossary changed conformance"));
        }
        canonical_catalog = canonical;
        &canonical_catalog
    };
    add_relationships(&mut root, &value.rels, value.conformance)?;
    validate_relationship_integrity(effective_catalog, &root, value.conformance)?;
    let staged_root_xml = root.blob_arc();
    if value.parts.len() > MAX_VALUES {
        return Err(invalid("glossary auxiliary part limit exceeded"));
    }
    let mut staged = HashMap::new();
    staged
        .try_reserve(value.parts.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary graph staging",
            source,
        })?;
    let mut total_bytes = 0usize;
    for auxiliary in &value.parts {
        total_bytes = total_bytes
            .checked_add(auxiliary.data().len())
            .ok_or_else(|| invalid("glossary auxiliary payload size overflow"))?;
        if total_bytes > MAX_GRAPH_BYTES {
            return Err(invalid(
                "glossary auxiliary graph exceeds the 256 MiB aggregate limit",
            ));
        }
        let uri = validate_physical_part(
            &auxiliary.name,
            &auxiliary.content_type,
            auxiliary.data().len(),
        )?;
        if uri == root_uri {
            return Err(invalid(format!(
                "glossary auxiliary part '{}' conflicts with the root",
                uri.as_str()
            )));
        }
        let mut part = BlobPart::new_shared(
            uri.clone(),
            auxiliary.content_type.clone(),
            auxiliary.shared_data(),
        );
        add_relationships(&mut part, &auxiliary.rels, value.conformance)?;
        if staged.insert(uri.clone(), part).is_some() {
            return Err(invalid(format!(
                "duplicate glossary auxiliary part '{uri}'"
            )));
        }
    }

    if let Some(current) = load_graph(package)?
        && graph_matches_catalog(&current, value, effective_catalog, Some(&staged_root_xml))?
    {
        return Ok(false);
    }

    validate_all_internal_targets(package)?;
    let owner = locate(package)?;
    let old_owned = if let Some(owner) = &owner {
        let owned = glossary_owned_parts(package, &owner.root, owner.conformance)?;
        validate_exclusive_ownership(package, owner, &owned)?;
        owned
    } else {
        HashSet::new()
    };

    let main = package.main_document_part()?.partname().clone();
    let mut candidate = package.clone();
    candidate.unsign();
    if let Some(owner) = &owner {
        candidate
            .get_part_mut(&main)?
            .rels_mut()
            .remove(&owner.relationship_id);
    }
    for uri in &old_owned {
        candidate.remove_part(uri);
    }
    for (_, part) in staged {
        candidate.try_add_part(Box::new(part))?;
    }
    candidate.try_add_part(Box::new(root))?;
    let generated_target = root_uri.relative_ref(main.base_uri());
    let owner_target = if value.owner_main.as_deref() == Some(main.as_str()) {
        value.owner_target.as_deref().unwrap_or(&generated_target)
    } else {
        &generated_target
    };
    let owner_relationships = candidate.get_part_mut(&main)?.rels_mut();
    let preserve_owner_id = value
        .owner_id
        .as_ref()
        .filter(|owner_id| owner_relationships.get(owner_id).is_none());
    if let Some(owner_id) = preserve_owner_id {
        owner_relationships.try_add_relationship(
            value.conformance.glossary_relationship().to_owned(),
            owner_target.to_owned(),
            owner_id.clone(),
            litchi_opc::TargetMode::Internal,
        )?;
    } else {
        owner_relationships.get_or_add(value.conformance.glossary_relationship(), owner_target);
    }

    validate_all_internal_targets(&candidate)?;
    let round_trip =
        load_graph(&candidate)?.ok_or_else(|| invalid("staged glossary graph is missing"))?;
    if !graph_matches_catalog(
        &round_trip,
        value,
        effective_catalog,
        Some(&staged_root_xml),
    )? {
        return Err(invalid("staged glossary graph did not round-trip"));
    }
    *package = candidate;
    Ok(true)
}

fn remove_graph(package: &mut OpcPackage) -> Result<Option<raw::Graph>> {
    let Some(owner) = locate(package)? else {
        return Ok(None);
    };
    let graph = load_graph(package)?.ok_or_else(|| invalid("missing glossary"))?;
    let owned = glossary_owned_parts(package, &owner.root, owner.conformance)?;
    validate_exclusive_ownership(package, &owner, &owned)?;
    validate_all_internal_targets(package)?;

    let mut candidate = package.clone();
    candidate.unsign();
    candidate
        .get_part_mut(&owner.main)?
        .rels_mut()
        .remove(&owner.relationship_id);
    for uri in owned {
        candidate.remove_part(&uri);
    }
    validate_all_internal_targets(&candidate)?;
    *package = candidate;
    Ok(Some(graph))
}

fn graph_matches_catalog(
    actual: &raw::Graph,
    expected: &raw::Graph,
    expected_catalog: &Catalog,
    expected_root_xml: Option<&Arc<Vec<u8>>>,
) -> Result<bool> {
    Ok(actual.catalog == *expected_catalog
        && actual.conformance == expected.conformance
        && actual.root_name == expected.root_name
        && expected_root_xml.is_none_or(|xml| {
            actual.root_xml.as_ref().is_some_and(|actual| {
                Arc::ptr_eq(actual, xml) || actual.as_slice() == xml.as_slice()
            })
        })
        && keyed_rels_match(&actual.rels, &expected.rels)?
        && keyed_parts_match(&actual.parts, &expected.parts)?)
}

fn keyed_rels_match(left: &[raw::Rel], right: &[raw::Rel]) -> Result<bool> {
    if left.len() != right.len() {
        return Ok(false);
    }
    let mut by_id = HashMap::new();
    by_id
        .try_reserve(right.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary relationship comparison",
            source,
        })?;
    for relationship in right {
        if by_id
            .insert(relationship.id.as_str(), relationship)
            .is_some()
        {
            return Ok(false);
        }
    }
    Ok(left
        .iter()
        .all(|relationship| by_id.get(relationship.id.as_str()) == Some(&relationship)))
}

fn keyed_parts_match(left: &[raw::Part], right: &[raw::Part]) -> Result<bool> {
    if left.len() != right.len() {
        return Ok(false);
    }
    let mut by_name = HashMap::new();
    by_name
        .try_reserve(right.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary part comparison",
            source,
        })?;
    for part in right {
        if by_name.insert(part.name.as_str(), part).is_some() {
            return Ok(false);
        }
    }
    for part in left {
        let Some(candidate) = by_name.get(part.name.as_str()) else {
            return Ok(false);
        };
        if part.content_type != candidate.content_type
            || (!Arc::ptr_eq(&part.data, &candidate.data) && part.data() != candidate.data())
            || !keyed_rels_match(&part.rels, &candidate.rels)?
        {
            return Ok(false);
        }
    }
    Ok(true)
}

fn is_signature_part(uri: &PackURI) -> bool {
    const PREFIX: &str = "/_xmlsignatures/";
    uri.as_str()
        .get(..PREFIX.len())
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case(PREFIX))
}

fn is_reserved_physical_part(uri: &PackURI) -> bool {
    let value = uri.as_str();
    if value == "/" || value.eq_ignore_ascii_case("/[Content_Types].xml") {
        return true;
    }
    let Some((directory, filename)) = value.rsplit_once('/') else {
        return false;
    };
    filename
        .get(filename.len().saturating_sub(5)..)
        .is_some_and(|suffix| suffix.eq_ignore_ascii_case(".rels"))
        && directory
            .rsplit('/')
            .next()
            .is_some_and(|segment| segment.eq_ignore_ascii_case("_rels"))
}

fn locate(package: &OpcPackage) -> Result<Option<Owner>> {
    let package_conformance = package_conformance(package)?;
    let main_part = package.main_document_part()?;
    if !matches!(
        main_part.content_type(),
        ct::WML_DOCUMENT_MAIN
            | ct::WML_TEMPLATE_MAIN
            | ct::WML_DOCUMENT_MACRO_MAIN
            | ct::WML_TEMPLATE_MACRO_MAIN
    ) {
        return Err(Error::ContentType {
            expected: format!(
                "{}, {}, {}, or {}",
                ct::WML_DOCUMENT_MAIN,
                ct::WML_TEMPLATE_MAIN,
                ct::WML_DOCUMENT_MACRO_MAIN,
                ct::WML_TEMPLATE_MACRO_MAIN,
            ),
            actual: main_part.content_type().to_owned(),
        });
    }
    let main = main_part.partname().clone();
    if package
        .rels()
        .iter()
        .any(|relationship| matches!(relationship.reltype(), REL | STRICT_REL))
    {
        return Err(invalid(
            "package root cannot source a glossary-document relationship",
        ));
    }

    let mut found = None;
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| matches!(relationship.reltype(), REL | STRICT_REL))
        {
            validate_relationship_metadata(
                relationship.r_id(),
                relationship.reltype(),
                relationship.target_ref(),
            )?;
            if part.partname() != &main {
                return Err(invalid(format!(
                    "glossary-document relationship has invalid source '{}'",
                    part.partname()
                )));
            }
            if found.is_some() {
                return Err(invalid("main document has multiple glossary relationships"));
            }
            if relationship.is_external() {
                return Err(invalid("glossary relationship cannot be external"));
            }
            let conformance = if relationship.reltype() == STRICT_REL {
                Conformance::Strict
            } else {
                Conformance::Transitional
            };
            if conformance != package_conformance {
                return Err(invalid(
                    "glossary relationship does not match package conformance",
                ));
            }
            let requested = relationship.target_partname()?;
            let target = package.get_part(&requested)?.partname().clone();
            let target_part = package.get_part(&target)?;
            if target_part.content_type() != CT {
                return Err(Error::ContentType {
                    expected: CT.to_owned(),
                    actual: target_part.content_type().to_owned(),
                });
            }
            found = Some(Owner {
                main: main.clone(),
                root: target,
                relationship_id: relationship.r_id().to_owned(),
                relationship_target: relationship.target_ref().to_owned(),
                conformance,
            });
        }
    }
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == CT)
    {
        match &found {
            Some(owner) if part.partname() == &owner.root => {},
            Some(_) => {
                return Err(invalid(format!(
                    "orphan glossary content-type part '{}' exists beside the owned root",
                    part.partname()
                )));
            },
            None => {
                return Err(invalid(format!(
                    "orphan glossary content-type part '{}' has no main-document relationship",
                    part.partname()
                )));
            },
        }
    }
    Ok(found)
}

fn package_conformance(package: &OpcPackage) -> Result<Conformance> {
    use litchi_opc::constants::relationship_type::{OFFICE_DOCUMENT, STRICT_OFFICE_DOCUMENT};
    let mut relationships = package.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            OFFICE_DOCUMENT | STRICT_OFFICE_DOCUMENT
        )
    });
    let relationship = relationships
        .next()
        .ok_or_else(|| invalid("main-document relationship is missing"))?;
    if relationships.next().is_some() {
        return Err(invalid("package has multiple main-document relationships"));
    }
    if relationship.is_external() {
        return Err(invalid("main-document relationship cannot be external"));
    }
    Ok(if relationship.reltype() == STRICT_OFFICE_DOCUMENT {
        Conformance::Strict
    } else {
        Conformance::Transitional
    })
}

fn validate_package_conformance(package: &OpcPackage, requested: Conformance) -> Result<()> {
    if package_conformance(package)? == requested {
        Ok(())
    } else {
        Err(invalid(
            "requested glossary conformance does not match the document package",
        ))
    }
}

fn add_relationships(
    part: &mut BlobPart,
    relationships: &[raw::Rel],
    conformance: Conformance,
) -> Result<()> {
    if relationships.len() > MAX_VALUES {
        return Err(invalid("glossary relationship limit exceeded"));
    }
    let mut ids = HashSet::new();
    for relationship in relationships {
        validate_relationship_metadata(&relationship.id, &relationship.kind, &relationship.target)?;
        if !ids.insert(relationship.id.clone()) {
            return Err(invalid("duplicate glossary relationship ID"));
        }
        relationship_kind(conformance, &relationship.kind).ok_or_else(|| {
            invalid(format!(
                "unsupported glossary relationship type '{}' for {conformance:?} conformance",
                relationship.kind
            ))
        })?;
        part.rels_mut().try_add_relationship(
            relationship.kind.clone(),
            relationship.target.clone(),
            relationship.id.clone(),
            if relationship.external {
                litchi_opc::TargetMode::External
            } else {
                litchi_opc::TargetMode::Internal
            },
        )?;
    }
    Ok(())
}

fn validate_relationship_metadata(id: &str, kind: &str, target: &str) -> Result<()> {
    bounded(id)?;
    bounded(kind)?;
    bounded(target)?;
    if !valid_ncname(id) {
        return Err(invalid(format!(
            "glossary relationship ID '{id}' is not an XML NCName"
        )));
    }
    if kind.is_empty()
        || kind
            .chars()
            .any(|character| character.is_whitespace() || character.is_control())
    {
        return Err(invalid("glossary relationship type is not a valid URI"));
    }
    if target.is_empty() || target.chars().any(char::is_control) {
        return Err(invalid(
            "glossary relationship target is empty or contains a control character",
        ));
    }
    Ok(())
}

fn validate_raw_graph_metadata(graph: &raw::Graph) -> Result<()> {
    if graph.parts.len() > MAX_VALUES {
        return Err(invalid("glossary auxiliary part limit exceeded"));
    }
    let mut relationship_count = 0usize;
    let mut metadata_bytes = 0usize;
    add_graph_metadata(&mut metadata_bytes, &graph.root_name)?;
    if let Some(owner_main) = &graph.owner_main {
        bounded(owner_main)?;
        PackURI::new(owner_main).map_err(Error::Uri)?;
        add_graph_metadata(&mut metadata_bytes, owner_main)?;
    }
    if let Some(owner_id) = &graph.owner_id {
        bounded(owner_id)?;
        if !valid_ncname(owner_id) {
            return Err(invalid(
                "glossary owner relationship ID is not an XML NCName",
            ));
        }
        add_graph_metadata(&mut metadata_bytes, owner_id)?;
    }
    if let Some(owner_target) = &graph.owner_target {
        bounded(owner_target)?;
        if owner_target.is_empty() || owner_target.chars().any(char::is_control) {
            return Err(invalid("glossary owner relationship target is invalid"));
        }
        add_graph_metadata(&mut metadata_bytes, owner_target)?;
    }
    validate_raw_relationship_set(&graph.rels, &mut relationship_count, &mut metadata_bytes)?;
    for part in &graph.parts {
        bounded(&part.name)?;
        bounded(&part.content_type)?;
        add_graph_metadata(&mut metadata_bytes, &part.name)?;
        add_graph_metadata(&mut metadata_bytes, &part.content_type)?;
        validate_raw_relationship_set(&part.rels, &mut relationship_count, &mut metadata_bytes)?;
    }
    Ok(())
}

fn validate_raw_relationship_set(
    relationships: &[raw::Rel],
    total_count: &mut usize,
    metadata_bytes: &mut usize,
) -> Result<()> {
    if relationships.len() > MAX_VALUES {
        return Err(invalid("glossary relationship limit exceeded"));
    }
    *total_count = total_count
        .checked_add(relationships.len())
        .ok_or_else(|| invalid("glossary relationship count overflow"))?;
    if *total_count > MAX_VALUES {
        return Err(invalid("glossary graph-wide relationship limit exceeded"));
    }
    for relationship in relationships {
        validate_relationship_metadata(&relationship.id, &relationship.kind, &relationship.target)?;
        add_graph_metadata(metadata_bytes, &relationship.id)?;
        add_graph_metadata(metadata_bytes, &relationship.kind)?;
        add_graph_metadata(metadata_bytes, &relationship.target)?;
    }
    Ok(())
}

fn add_graph_metadata(total: &mut usize, value: &str) -> Result<()> {
    *total = total
        .checked_add(value.len())
        .ok_or_else(|| invalid("glossary graph metadata size overflow"))?;
    if *total > MAX_GRAPH_METADATA_BYTES {
        return Err(invalid(
            "glossary graph metadata exceeds the 32 MiB aggregate limit",
        ));
    }
    Ok(())
}

fn relationship_kind(conformance: Conformance, value: &str) -> Option<&str> {
    if value == STYLES_EFFECTS_REL {
        return Some("stylesWithEffects");
    }
    if value == CUSTOMIZATIONS_REL {
        return Some("keyMapCustomizations");
    }
    if value == ATTACHED_TOOLBARS_REL {
        return Some("attachedToolbars");
    }
    if value == DIAGRAM_DRAWING_REL {
        return Some("diagramDrawing");
    }
    if matches!(value, CHART_STYLE_REL | CHART_STYLE_REL_2012) {
        return Some("chartStyle");
    }
    if matches!(value, CHART_COLOR_STYLE_REL | CHART_COLOR_STYLE_REL_2012) {
        return Some("chartColorStyle");
    }
    if value == ACTIVE_X_BINARY_REL {
        return Some("activeXControlBinary");
    }
    value
        .strip_prefix(conformance.relationships())
        .and_then(|kind| kind.strip_prefix('/'))
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum GraphRole {
    Glossary,
    RichStory,
    Settings,
    FontTable,
    Numbering,
    WebSettings,
    Chart,
    ChartDrawing,
    DiagramData,
    DiagramLayout,
    EmbeddedObject,
    EmbeddedPackage,
    Control,
    ActiveX,
    CustomXml,
    Customizations,
    Leaf,
}

#[derive(Clone, Copy)]
enum EdgeMode {
    Internal,
    External,
    Either,
}

#[derive(Clone, Copy)]
enum TargetProfile {
    Exact(&'static str),
    Image,
    Video,
    Xml,
    Font,
    Any,
}

#[derive(Clone, Copy)]
struct EdgeSpec {
    mode: EdgeMode,
    target: TargetProfile,
    role: Option<GraphRole>,
    owned: bool,
}

fn edge_spec(conformance: Conformance, role: GraphRole, value: &str) -> Result<EdgeSpec> {
    let kind = relationship_kind(conformance, value).ok_or_else(|| {
        invalid(format!(
            "unsupported glossary relationship type '{value}' for {conformance:?} conformance"
        ))
    })?;
    let internal = |target, role| EdgeSpec {
        mode: EdgeMode::Internal,
        target,
        role: Some(role),
        owned: true,
    };
    let either = |target, role| EdgeSpec {
        mode: EdgeMode::Either,
        target,
        role: Some(role),
        owned: true,
    };
    let reference = EdgeSpec {
        mode: EdgeMode::Either,
        target: TargetProfile::Any,
        role: None,
        owned: false,
    };
    let external = EdgeSpec {
        mode: EdgeMode::External,
        target: TargetProfile::Any,
        role: None,
        owned: false,
    };
    let spec = match (role, kind) {
        (GraphRole::Glossary, "comments") => {
            internal(TargetProfile::Exact(ct::WML_COMMENTS), GraphRole::RichStory)
        },
        (GraphRole::Glossary, "settings") => {
            internal(TargetProfile::Exact(ct::WML_SETTINGS), GraphRole::Settings)
        },
        (GraphRole::Glossary, "endnotes") => {
            internal(TargetProfile::Exact(ct::WML_ENDNOTES), GraphRole::RichStory)
        },
        (GraphRole::Glossary, "fontTable") => internal(
            TargetProfile::Exact(ct::WML_FONT_TABLE),
            GraphRole::FontTable,
        ),
        (GraphRole::Glossary, "footnotes") => internal(
            TargetProfile::Exact(ct::WML_FOOTNOTES),
            GraphRole::RichStory,
        ),
        (GraphRole::Glossary, "numbering") => internal(
            TargetProfile::Exact(ct::WML_NUMBERING),
            GraphRole::Numbering,
        ),
        (GraphRole::Glossary, "styles") => {
            internal(TargetProfile::Exact(ct::WML_STYLES), GraphRole::Leaf)
        },
        (GraphRole::Glossary, "webSettings") => internal(
            TargetProfile::Exact(ct::WML_WEB_SETTINGS),
            GraphRole::WebSettings,
        ),
        (GraphRole::Glossary, "aFChunk") => internal(TargetProfile::Any, GraphRole::Leaf),
        (GraphRole::Glossary, "chart") => {
            internal(TargetProfile::Exact(ct::DML_CHART), GraphRole::Chart)
        },
        (GraphRole::Glossary, "customXml") => internal(TargetProfile::Xml, GraphRole::CustomXml),
        (GraphRole::Glossary, "diagramColors") => internal(
            TargetProfile::Exact(ct::DML_DIAGRAM_COLORS),
            GraphRole::Leaf,
        ),
        (GraphRole::Glossary, "diagramData") => internal(
            TargetProfile::Exact(ct::DML_DIAGRAM_DATA),
            GraphRole::DiagramData,
        ),
        (GraphRole::Glossary, "diagramLayout") => internal(
            TargetProfile::Exact(ct::DML_DIAGRAM_LAYOUT),
            GraphRole::DiagramLayout,
        ),
        (GraphRole::Glossary, "diagramQuickStyle") => {
            internal(TargetProfile::Exact(ct::DML_DIAGRAM_STYLE), GraphRole::Leaf)
        },
        (GraphRole::Glossary, "control") => internal(TargetProfile::Any, GraphRole::Control),
        (GraphRole::Glossary, "oleObject") => either(TargetProfile::Any, GraphRole::EmbeddedObject),
        (GraphRole::Glossary, "package") => either(TargetProfile::Any, GraphRole::EmbeddedPackage),
        (GraphRole::Glossary, "footer") => {
            internal(TargetProfile::Exact(ct::WML_FOOTER), GraphRole::RichStory)
        },
        (GraphRole::Glossary, "header") => {
            internal(TargetProfile::Exact(ct::WML_HEADER), GraphRole::RichStory)
        },
        (GraphRole::Glossary, "hyperlink") => reference,
        (GraphRole::Glossary, "image") => either(TargetProfile::Image, GraphRole::Leaf),
        (GraphRole::Glossary, "printerSettings") => {
            internal(TargetProfile::Exact(PRINTER_SETTINGS_CT), GraphRole::Leaf)
        },
        (GraphRole::Glossary, "video") => either(TargetProfile::Video, GraphRole::Leaf),
        (GraphRole::Glossary, "stylesWithEffects") => {
            internal(TargetProfile::Exact(STYLES_EFFECTS_CT), GraphRole::Leaf)
        },
        (GraphRole::Glossary, "keyMapCustomizations") => internal(
            TargetProfile::Exact(CUSTOMIZATIONS_CT),
            GraphRole::Customizations,
        ),
        (GraphRole::RichStory, "aFChunk") => internal(TargetProfile::Any, GraphRole::Leaf),
        (GraphRole::RichStory, "control") => internal(TargetProfile::Any, GraphRole::Control),
        (GraphRole::RichStory, "oleObject") => {
            either(TargetProfile::Any, GraphRole::EmbeddedObject)
        },
        (GraphRole::RichStory, "package") => either(TargetProfile::Any, GraphRole::EmbeddedPackage),
        (GraphRole::RichStory, "hyperlink") => reference,
        (GraphRole::RichStory, "image") => either(TargetProfile::Image, GraphRole::Leaf),
        (GraphRole::RichStory, "video") => either(TargetProfile::Video, GraphRole::Leaf),
        (GraphRole::RichStory, "chart") => {
            internal(TargetProfile::Exact(ct::DML_CHART), GraphRole::Chart)
        },
        (GraphRole::RichStory, "customXml") => internal(TargetProfile::Xml, GraphRole::CustomXml),
        (GraphRole::RichStory, "diagramColors") => internal(
            TargetProfile::Exact(ct::DML_DIAGRAM_COLORS),
            GraphRole::Leaf,
        ),
        (GraphRole::RichStory, "diagramData") => internal(
            TargetProfile::Exact(ct::DML_DIAGRAM_DATA),
            GraphRole::DiagramData,
        ),
        (GraphRole::RichStory, "diagramLayout") => internal(
            TargetProfile::Exact(ct::DML_DIAGRAM_LAYOUT),
            GraphRole::DiagramLayout,
        ),
        (GraphRole::RichStory, "diagramQuickStyle") => {
            internal(TargetProfile::Exact(ct::DML_DIAGRAM_STYLE), GraphRole::Leaf)
        },
        (
            GraphRole::Settings,
            "attachedTemplate" | "mailMergeSource" | "mailMergeHeaderSource" | "transform",
        ) => external,
        (GraphRole::Settings, "recipientData") => {
            internal(TargetProfile::Exact(RECIPIENT_DATA_CT), GraphRole::Leaf)
        },
        (GraphRole::FontTable, "font") => internal(TargetProfile::Font, GraphRole::Leaf),
        (GraphRole::Numbering, "image") => either(TargetProfile::Image, GraphRole::Leaf),
        (GraphRole::WebSettings, "frame") => external,
        (GraphRole::Chart, "chartUserShapes") => internal(
            TargetProfile::Exact(ct::DML_CHARTSHAPES),
            GraphRole::ChartDrawing,
        ),
        (GraphRole::Chart, "chartStyle") => {
            internal(TargetProfile::Exact(CHART_STYLE_CT), GraphRole::Leaf)
        },
        (GraphRole::Chart, "chartColorStyle") => {
            internal(TargetProfile::Exact(CHART_COLOR_STYLE_CT), GraphRole::Leaf)
        },
        (GraphRole::Chart, "themeOverride") => internal(
            TargetProfile::Exact(ct::OFC_THEME_OVERRIDE),
            GraphRole::Leaf,
        ),
        (GraphRole::Chart, "package") => either(TargetProfile::Any, GraphRole::EmbeddedPackage),
        (GraphRole::ChartDrawing, "image") => either(TargetProfile::Image, GraphRole::Leaf),
        (GraphRole::ChartDrawing, "chart") => {
            internal(TargetProfile::Exact(ct::DML_CHART), GraphRole::Chart)
        },
        (GraphRole::ChartDrawing, "customXml") => {
            internal(TargetProfile::Xml, GraphRole::CustomXml)
        },
        (GraphRole::DiagramData, "image") | (GraphRole::DiagramLayout, "image") => {
            either(TargetProfile::Image, GraphRole::Leaf)
        },
        (GraphRole::DiagramData, "diagramDrawing") => internal(
            TargetProfile::Exact(ct::DML_DIAGRAM_DRAWING),
            GraphRole::Leaf,
        ),
        (GraphRole::DiagramData, "hyperlink") => reference,
        (GraphRole::EmbeddedObject, "hyperlink") | (GraphRole::EmbeddedPackage, "hyperlink") => {
            reference
        },
        (GraphRole::ActiveX, "activeXControlBinary") => {
            internal(TargetProfile::Exact(ACTIVE_X_BINARY_CT), GraphRole::Leaf)
        },
        (GraphRole::CustomXml, "customXmlProps") => internal(
            TargetProfile::Exact(litchi_ooxml_common::custom_xml::PROPS_CONTENT_TYPE),
            GraphRole::Leaf,
        ),
        (GraphRole::Customizations, "attachedToolbars") => {
            internal(TargetProfile::Exact(ATTACHED_TOOLBARS_CT), GraphRole::Leaf)
        },
        _ => {
            return Err(invalid(format!(
                "relationship type '{value}' is invalid for glossary role {role:?}"
            )));
        },
    };
    Ok(spec)
}

fn validate_edge_mode(kind: &str, mode: EdgeMode, external: bool) -> Result<()> {
    let valid = match mode {
        EdgeMode::Internal => !external,
        EdgeMode::External => external,
        EdgeMode::Either => true,
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(format!(
            "glossary relationship kind '{kind}' has an invalid target mode"
        )))
    }
}

fn validate_target_profile(kind: &str, profile: TargetProfile, content_type: &str) -> Result<()> {
    ContentType::new(content_type.to_owned())?;
    let media_type = content_type.split(';').next().unwrap_or_default();
    let valid = match profile {
        TargetProfile::Exact(expected) => media_type.eq_ignore_ascii_case(expected),
        TargetProfile::Image => starts_with_ascii_case_insensitive(media_type, "image/"),
        TargetProfile::Video => starts_with_ascii_case_insensitive(media_type, "video/"),
        TargetProfile::Xml => {
            media_type.eq_ignore_ascii_case("application/xml")
                || media_type.eq_ignore_ascii_case("text/xml")
                || ends_with_ascii_case_insensitive(media_type, "+xml")
        },
        TargetProfile::Font => [FONT_DATA_CT, FONT_TTF_CT, OBFUSCATED_FONT_CT]
            .iter()
            .any(|expected| media_type.eq_ignore_ascii_case(expected)),
        TargetProfile::Any => true,
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(format!(
            "glossary relationship kind '{kind}' cannot target content type '{content_type}'"
        )))
    }
}

fn target_graph_role(role: GraphRole, content_type: &str) -> GraphRole {
    if role == GraphRole::Control
        && content_type
            .split(';')
            .next()
            .is_some_and(|value| value.eq_ignore_ascii_case(ACTIVE_X_DESCRIPTOR_CT))
    {
        GraphRole::ActiveX
    } else {
        role
    }
}

fn starts_with_ascii_case_insensitive(value: &str, prefix: &str) -> bool {
    value
        .get(..prefix.len())
        .is_some_and(|candidate| candidate.eq_ignore_ascii_case(prefix))
}

fn ends_with_ascii_case_insensitive(value: &str, suffix: &str) -> bool {
    value
        .get(value.len().saturating_sub(suffix.len())..)
        .is_some_and(|candidate| candidate.eq_ignore_ascii_case(suffix))
}

fn validate_relationship_integrity(
    document: &Catalog,
    part: &dyn Part,
    conformance: Conformance,
) -> Result<()> {
    for id in catalog_relationship_references(document, conformance)? {
        if part.rels().get(id).is_none() {
            return Err(invalid(format!(
                "glossary XML references missing relationship '{id}'"
            )));
        }
    }
    for relationship in part.rels().iter() {
        relationship_kind(conformance, relationship.reltype()).ok_or_else(|| {
            invalid(format!(
                "unsupported glossary relationship type '{}' for {conformance:?} conformance",
                relationship.reltype()
            ))
        })?;
    }
    Ok(())
}

fn catalog_relationship_references(
    document: &Catalog,
    conformance: Conformance,
) -> Result<HashSet<&str>> {
    let mut reference_count = document.background_refs.len();
    for entry in &document.entries {
        let refs = entry
            .producer
            .as_deref()
            .filter(|producer| producer.conformance == conformance)
            .map_or(entry.refs.as_ref(), |producer| producer.refs.as_ref());
        reference_count = reference_count
            .checked_add(refs.len())
            .ok_or_else(|| invalid("glossary relationship reference count overflow"))?;
    }
    let mut referenced = HashSet::new();
    referenced
        .try_reserve(reference_count)
        .map_err(|source| Error::Allocation {
            resource: "glossary relationship reference index",
            source,
        })?;
    referenced.extend(document.background_refs.iter().map(String::as_str));
    for entry in &document.entries {
        let refs = entry
            .producer
            .as_deref()
            .filter(|producer| producer.conformance == conformance)
            .map_or(entry.refs.as_ref(), |producer| producer.refs.as_ref());
        referenced.extend(refs.iter().map(String::as_str));
    }
    Ok(referenced)
}

fn collect_relationship_references(node: &Node, output: &mut HashSet<String>) {
    for attribute in &node.attrs {
        if matches!(attribute.ns.as_ref(), R | RS) {
            output.insert(attribute.v.clone());
        }
    }
    for content in &node.content {
        if let Content::Node(child) = content {
            collect_relationship_references(child, output);
        }
    }
}

fn relationship_references(node: &Node) -> Result<Arc<[String]>> {
    let mut refs = HashSet::new();
    collect_relationship_references(node, &mut refs);
    let mut refs = refs.into_iter().collect::<Vec<_>>();
    refs.sort();
    Ok(Arc::from(refs))
}

fn glossary_owned_parts(
    package: &OpcPackage,
    root: &PackURI,
    conformance: Conformance,
) -> Result<HashSet<PackURI>> {
    if is_signature_part(root) || is_reserved_physical_part(root) {
        return Err(invalid(
            "glossary root cannot use reserved OPC package infrastructure",
        ));
    }
    let root = package.get_part(root)?.partname().clone();
    let mut owned = HashSet::from([root.clone()]);
    let mut roles = HashMap::from([(root.clone(), GraphRole::Glossary)]);
    let mut queue = VecDeque::from([(root, GraphRole::Glossary)]);
    let mut relationship_count = 0usize;
    let mut metadata_bytes = 0usize;
    while let Some((uri, role)) = queue.pop_front() {
        let part = package.get_part(&uri)?;
        validate_physical_part(uri.as_str(), part.content_type(), part.blob().len())?;
        add_graph_metadata(&mut metadata_bytes, uri.as_str())?;
        add_graph_metadata(&mut metadata_bytes, part.content_type())?;
        let part_relationship_count = part.rels().iter().count();
        if part_relationship_count > MAX_VALUES {
            return Err(invalid("glossary relationship limit exceeded"));
        }
        relationship_count = relationship_count
            .checked_add(part_relationship_count)
            .ok_or_else(|| invalid("glossary relationship count overflow"))?;
        if relationship_count > MAX_VALUES {
            return Err(invalid("glossary graph-wide relationship limit exceeded"));
        }
        for relationship in part.rels().iter() {
            validate_relationship_metadata(
                relationship.r_id(),
                relationship.reltype(),
                relationship.target_ref(),
            )?;
            add_graph_metadata(&mut metadata_bytes, relationship.r_id())?;
            add_graph_metadata(&mut metadata_bytes, relationship.reltype())?;
            add_graph_metadata(&mut metadata_bytes, relationship.target_ref())?;
        }
        for relationship in part.rels().iter() {
            let kind = relationship_kind(conformance, relationship.reltype()).ok_or_else(|| {
                invalid(format!(
                    "unsupported glossary relationship type '{}' for {conformance:?} conformance",
                    relationship.reltype()
                ))
            })?;
            let spec = edge_spec(conformance, role, relationship.reltype())?;
            validate_edge_mode(kind, spec.mode, relationship.is_external())?;
            if relationship.is_external() {
                continue;
            }
            let requested = relationship.target_partname()?;
            let target = package.get_part(&requested)?.partname().clone();
            if is_signature_part(&target) || is_reserved_physical_part(&target) {
                return Err(invalid(format!(
                    "glossary relationship '{}' targets reserved OPC package infrastructure",
                    relationship.r_id()
                )));
            }
            let target_part = package.get_part(&target)?;
            validate_target_profile(kind, spec.target, target_part.content_type())?;
            if !spec.owned {
                continue;
            }
            let target_role = spec
                .role
                .ok_or_else(|| invalid("owned glossary relationship is missing a target role"))?;
            let target_role = target_graph_role(target_role, target_part.content_type());
            if let Some(existing_role) = roles.get(&target) {
                if *existing_role != target_role {
                    return Err(invalid(format!(
                        "glossary-owned part '{target}' has conflicting graph roles"
                    )));
                }
                continue;
            }
            roles.insert(target.clone(), target_role);
            if owned.insert(target.clone()) {
                if owned.len() > MAX_VALUES + 1 {
                    return Err(invalid("glossary owned-part limit exceeded"));
                }
                queue.push_back((target, target_role));
            }
        }
    }
    Ok(owned)
}

fn validate_exclusive_ownership(
    package: &OpcPackage,
    owner: &Owner,
    owned: &HashSet<PackURI>,
) -> Result<()> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        let target = package
            .get_part(&relationship.target_partname()?)?
            .partname();
        if owned.contains(target) {
            return Err(invalid(format!(
                "glossary-owned part '{target}' has an inbound package-root relationship"
            )));
        }
    }
    for source in package.iter_parts() {
        for relationship in source
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            let target = package
                .get_part(&relationship.target_partname()?)?
                .partname();
            if owned.contains(target)
                && !owned.contains(source.partname())
                && !(source.partname() == &owner.main
                    && relationship.r_id() == owner.relationship_id)
            {
                return Err(invalid(format!(
                    "glossary-owned part '{target}' is shared by '{}'",
                    source.partname()
                )));
            }
        }
    }
    Ok(())
}

fn validate_all_internal_targets(package: &OpcPackage) -> Result<()> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        package.get_part(&relationship.target_partname()?)?;
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            package.get_part(&relationship.target_partname()?)?;
        }
    }
    Ok(())
}

#[derive(Clone, Debug, Default)]
struct NamespaceFrame {
    parent: Option<Arc<NamespaceFrame>>,
    local: Vec<(String, Arc<str>)>,
    declaration_count: usize,
}

#[derive(Clone, Debug)]
struct Attr {
    q: String,
    ns: Arc<str>,
    l: String,
    v: String,
}
#[derive(Clone, Debug)]
enum Content {
    Node(Node),
    Text(String),
    CData(String),
    Comment(String),
}
#[derive(Clone, Debug)]
struct Node {
    q: String,
    ns: Arc<str>,
    l: String,
    attrs: Vec<Attr>,
    bindings: Arc<NamespaceFrame>,
    content: Vec<Content>,
}

#[derive(Default)]
struct DomBudget {
    bytes: usize,
    nodes: usize,
    attributes: usize,
    contents: usize,
    tokens: usize,
}

impl DomBudget {
    fn charge(&mut self, bytes: usize) -> Result<()> {
        self.bytes = self
            .bytes
            .checked_add(bytes)
            .ok_or_else(|| invalid("glossary DOM allocation budget overflow"))?;
        if self.bytes > MAX_DOM_BYTES {
            return Err(invalid(
                "glossary DOM exceeds the 64 MiB owned-allocation budget",
            ));
        }
        Ok(())
    }

    fn token(&mut self) -> Result<()> {
        self.tokens = self
            .tokens
            .checked_add(1)
            .ok_or_else(|| invalid("glossary XML token count overflow"))?;
        if self.tokens > MAX_DOM_TOKENS {
            return Err(invalid("glossary XML token limit exceeded"));
        }
        Ok(())
    }

    fn node(&mut self) -> Result<()> {
        self.nodes = self
            .nodes
            .checked_add(1)
            .ok_or_else(|| invalid("glossary XML node count overflow"))?;
        if self.nodes > MAX_NODES {
            return Err(invalid("glossary XML node limit exceeded"));
        }
        self.charge(std::mem::size_of::<Node>() + std::mem::size_of::<NamespaceFrame>())
    }

    fn node_names(&mut self, qualified: &str, local: &str) -> Result<()> {
        self.charge(
            qualified
                .len()
                .checked_add(local.len())
                .ok_or_else(|| invalid("glossary node-name budget overflow"))?,
        )
    }

    fn binding(&mut self, prefix: &str, namespace: &str) -> Result<()> {
        self.charge(
            std::mem::size_of::<(String, Arc<str>)>()
                .checked_add(prefix.len())
                .and_then(|bytes| bytes.checked_add(namespace.len()))
                .ok_or_else(|| invalid("glossary namespace budget overflow"))?,
        )
    }

    fn attribute(&mut self, qualified: &str, local: &str, value: &str) -> Result<()> {
        self.attributes = self
            .attributes
            .checked_add(1)
            .ok_or_else(|| invalid("glossary XML attribute count overflow"))?;
        if self.attributes > MAX_DOM_ATTRIBUTES {
            return Err(invalid("glossary XML aggregate attribute limit exceeded"));
        }
        self.charge(
            std::mem::size_of::<Attr>()
                .checked_add(qualified.len())
                .and_then(|bytes| bytes.checked_add(local.len()))
                .and_then(|bytes| bytes.checked_add(value.len()))
                .ok_or_else(|| invalid("glossary attribute budget overflow"))?,
        )
    }

    fn content(&mut self, content: &Content) -> Result<()> {
        self.contents = self
            .contents
            .checked_add(1)
            .ok_or_else(|| invalid("glossary XML content count overflow"))?;
        if self.contents > MAX_DOM_CONTENT {
            return Err(invalid("glossary XML content token limit exceeded"));
        }
        let owned = match content {
            Content::Node(_) => 0,
            Content::Text(value) | Content::CData(value) | Content::Comment(value) => value.len(),
        };
        self.charge(
            std::mem::size_of::<Content>()
                .checked_add(owned)
                .ok_or_else(|| invalid("glossary content budget overflow"))?,
        )
    }
}

#[derive(Default)]
struct NamespaceResolver {
    active: HashMap<String, Vec<Arc<str>>>,
}

impl NamespaceResolver {
    fn push(&mut self, bindings: &[(String, Arc<str>)]) -> Result<()> {
        self.active
            .try_reserve(bindings.len())
            .map_err(|source| Error::Allocation {
                resource: "glossary active namespace resolver",
                source,
            })?;
        for (prefix, namespace) in bindings {
            if let Some(values) = self.active.get_mut(prefix) {
                values.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "glossary active namespace stack",
                    source,
                })?;
                values.push(Arc::clone(namespace));
            } else {
                let mut values = Vec::new();
                values
                    .try_reserve_exact(1)
                    .map_err(|source| Error::Allocation {
                        resource: "glossary active namespace stack",
                        source,
                    })?;
                values.push(Arc::clone(namespace));
                self.active.insert(prefix.clone(), values);
            }
        }
        Ok(())
    }

    fn pop(&mut self, bindings: &[(String, Arc<str>)]) -> Result<()> {
        for (prefix, _) in bindings.iter().rev() {
            let remove = {
                let values = self
                    .active
                    .get_mut(prefix)
                    .ok_or_else(|| invalid("glossary namespace stack is inconsistent"))?;
                values
                    .pop()
                    .ok_or_else(|| invalid("glossary namespace stack is empty"))?;
                values.is_empty()
            };
            if remove {
                self.active.remove(prefix);
            }
        }
        Ok(())
    }

    fn resolve(&self, prefix: &str) -> Result<Arc<str>> {
        if prefix == "xml" {
            return Ok(Arc::from("http://www.w3.org/XML/1998/namespace"));
        }
        self.active
            .get(prefix)
            .and_then(|values| values.last())
            .map(Arc::clone)
            .ok_or_else(|| invalid(format!("unbound prefix '{prefix}'")))
    }
}

fn validate_xml_value(value: &str, context: &str) -> Result<()> {
    if value.chars().all(xml_char) {
        Ok(())
    } else {
        Err(invalid(format!(
            "glossary XML {context} contains a character forbidden by XML 1.0"
        )))
    }
}

fn parse_dom(xml: &[u8]) -> Result<Node> {
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut rd = Reader::from_reader(xml);
    let mut budget = DomBudget::default();
    let mut resolver = NamespaceResolver::default();
    let mut stack: Vec<Node> = Vec::new();
    stack
        .try_reserve_exact(MAX_DEPTH)
        .map_err(|source| Error::Allocation {
            resource: "glossary XML stack",
            source,
        })?;
    let mut root = None;
    loop {
        let d = rd.decoder();
        let event = rd.read_event().map_err(xml_error)?;
        if !matches!(&event, Event::Eof) {
            budget.token()?;
        }
        match event {
            Event::Start(e) => {
                budget.node()?;
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid("glossary XML resource limit exceeded"));
                }
                stack.push(make(&e, d, &stack, &mut resolver, &mut budget)?);
            },
            Event::Empty(e) => {
                budget.node()?;
                let n = make(&e, d, &stack, &mut resolver, &mut budget)?;
                resolver.pop(&n.bindings.local)?;
                attach(&mut stack, &mut root, n, &mut budget)?;
            },
            Event::End(_) => {
                let n = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected glossary closing element"))?;
                resolver.pop(&n.bindings.local)?;
                attach(&mut stack, &mut root, n, &mut budget)?;
            },
            Event::Text(t) => {
                let v = t.decode().map_err(xml_error)?.into_owned();
                validate_xml_value(&v, "text")?;
                if let Some(n) = stack.last_mut() {
                    push_content(n, Content::Text(v), &mut budget)?
                } else if !v.trim().is_empty() {
                    return Err(invalid("text outside glossary root"));
                }
            },
            Event::CData(t) => {
                let value = t.decode().map_err(xml_error)?.into_owned();
                validate_xml_value(&value, "CDATA")?;
                if let Some(n) = stack.last_mut() {
                    push_content(n, Content::CData(value), &mut budget)?
                } else {
                    return Err(invalid("CDATA outside glossary root"));
                }
            },
            Event::Comment(t) => {
                let value = t.decode().map_err(xml_error)?.into_owned();
                validate_xml_value(&value, "comment")?;
                if value.contains("--") || value.ends_with('-') {
                    return Err(invalid("glossary XML comment has an invalid lexical form"));
                }
                if let Some(n) = stack.last_mut() {
                    push_content(n, Content::Comment(value), &mut budget)?
                }
            },
            Event::GeneralRef(t) => {
                let value =
                    litchi_ooxml_common::xml::decode_xml_reference(&t).map_err(xml_error)?;
                validate_xml_value(&value, "character reference")?;
                if let Some(n) = stack.last_mut() {
                    push_content(n, Content::Text(value), &mut budget)?
                } else {
                    return Err(invalid("entity outside glossary root"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Decl(_) => {},
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated glossary XML"));
    }
    root.ok_or_else(|| invalid("missing glossary root"))
}
fn make(
    e: &BytesStart<'_>,
    d: Decoder,
    stack: &[Node],
    resolver: &mut NamespaceResolver,
    budget: &mut DomBudget,
) -> Result<Node> {
    let q = std::str::from_utf8(e.name().as_ref())
        .map_err(xml_error)?
        .to_string();
    validate_xml_value(&q, "qualified name")?;
    let parent = stack.last().map(|node| Arc::clone(&node.bindings));
    let mut raw = Vec::new();
    for a in e.attributes().with_checks(true) {
        if raw.len() >= MAX_VALUES {
            return Err(invalid("glossary XML attribute limit exceeded"));
        }
        raw.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "glossary XML attributes",
            source,
        })?;
        let a = a.map_err(xml_error)?;
        let key = std::str::from_utf8(a.key.as_ref())
            .map_err(xml_error)?
            .to_string();
        let value = a
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, d)
            .map_err(xml_error)?
            .into_owned();
        validate_xml_value(&key, "attribute name")?;
        validate_xml_value(&value, "attribute value")?;
        raw.push((key, value));
    }
    let mut local_bindings: Vec<(String, Arc<str>)> = Vec::new();
    for (k, v) in &raw {
        if k == "xmlns" || k.starts_with("xmlns:") {
            let key = k.strip_prefix("xmlns:").unwrap_or("").to_string();
            local_bindings
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "glossary namespace bindings",
                    source,
                })?;
            local_bindings.push((key, Arc::from(v.as_str())));
        }
    }
    let declaration_count = parent
        .as_ref()
        .map_or(0, |frame| frame.declaration_count)
        .checked_add(local_bindings.len())
        .ok_or_else(|| invalid("glossary namespace binding count overflow"))?;
    if declaration_count > MAX_VALUES {
        return Err(invalid("glossary namespace binding limit exceeded"));
    }
    let bindings = Arc::new(NamespaceFrame {
        parent,
        local: local_bindings,
        declaration_count,
    });
    for (prefix, namespace) in &bindings.local {
        budget.binding(prefix, namespace)?;
    }
    resolver.push(&bindings.local)?;
    let (pr, lo) = split(&q)?;
    let local = lo.to_string();
    budget.node_names(&q, &local)?;
    let ns = resolver.resolve(pr)?;
    let mut attrs = Vec::new();
    for (q, v) in raw {
        if q == "xmlns" || q.starts_with("xmlns:") {
            continue;
        }
        let (pr, lo) = split(&q)?;
        let ans = if pr.is_empty() {
            Arc::from("")
        } else {
            resolver.resolve(pr)?
        };
        let local = lo.to_string();
        budget.attribute(&q, &local, &v)?;
        attrs.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "glossary typed attributes",
            source,
        })?;
        attrs.push(Attr {
            q,
            ns: ans,
            l: local,
            v,
        });
    }
    Ok(Node {
        q,
        ns,
        l: local,
        attrs,
        bindings,
        content: Vec::new(),
    })
}
fn attach(
    stack: &mut [Node],
    root: &mut Option<Node>,
    n: Node,
    budget: &mut DomBudget,
) -> Result<()> {
    if let Some(p) = stack.last_mut() {
        push_content(p, Content::Node(n), budget)?
    } else if root.replace(n).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

fn push_content(node: &mut Node, content: Content, budget: &mut DomBudget) -> Result<()> {
    budget.content(&content)?;
    node.content
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "glossary XML content",
            source,
        })?;
    node.content.push(content);
    Ok(())
}

fn project(n: &Node) -> Result<Catalog> {
    expect(n, "glossaryDocument")?;
    noattrs(n)?;
    let c = kids(n)?;
    let mut out = Catalog::default();
    let mut semantic_bytes = 0usize;
    let mut index = 0;
    if let Some(background) = c.first().filter(|node| node.l == "background") {
        expect(background, "background")?;
        let xml = node_xml(background, false)?;
        add_semantic_xml_bytes(&mut semantic_bytes, xml.len())?;
        out.background = Some(xml);
        index = 1;
    }
    if index < c.len() {
        let parts = c
            .get(index)
            .copied()
            .ok_or_else(|| invalid("glossary child index is out of bounds"))?;
        expect(parts, "docParts")?;
        noattrs(parts)?;
        for p in kids(parts)? {
            if out.entries.len() >= MAX_PARTS {
                return Err(invalid("glossary entry limit exceeded"));
            }
            out.entries.push(parse_entry(p, &mut semantic_bytes)?);
        }
        index += 1;
    }
    if index != c.len() {
        return Err(invalid("unexpected glossary root child"));
    }
    validate_catalog_fields(&out)?;
    Ok(out)
}

fn add_semantic_xml_bytes(total: &mut usize, bytes: usize) -> Result<()> {
    *total = total
        .checked_add(bytes)
        .ok_or_else(|| invalid("glossary semantic XML size overflow"))?;
    if *total > MAX {
        return Err(invalid(
            "glossary semantic XML exceeds the 32 MiB aggregate limit",
        ));
    }
    Ok(())
}

fn extract_producer_entries(
    root: Node,
    conformance: Conformance,
) -> Result<Option<Vec<ProducerEntry>>> {
    let mut inherited_mce = Vec::new();
    merge_mce_attributes(&mut inherited_mce, &root)?;
    let mut doc_parts = None;
    for content in root.content {
        match content {
            Content::Node(node)
                if node.ns.as_ref() == conformance.word() && node.l == "docParts" =>
            {
                if doc_parts.replace(node).is_some() {
                    return Ok(None);
                }
            },
            Content::Text(text) if text.trim().is_empty() => {},
            Content::Comment(_) | Content::Node(_) => {},
            Content::Text(_) | Content::CData(_) => return Ok(None),
        }
    }
    let Some(doc_parts) = doc_parts else {
        return Ok(None);
    };
    merge_mce_attributes(&mut inherited_mce, &doc_parts)?;
    let mut producer_nodes = Vec::new();
    for content in doc_parts.content {
        match content {
            Content::Node(node)
                if node.ns.as_ref() == conformance.word() && node.l == "docPart" =>
            {
                if producer_nodes.len() >= MAX_PARTS {
                    return Err(invalid("glossary producer entry limit exceeded"));
                }
                producer_nodes
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "glossary producer nodes",
                        source,
                    })?;
                producer_nodes.push(node);
            },
            Content::Text(text) if text.trim().is_empty() => {},
            Content::Comment(_) => {},
            Content::Node(_) | Content::Text(_) | Content::CData(_) => return Ok(None),
        }
    }
    let mut snapshots = Vec::new();
    snapshots
        .try_reserve_exact(producer_nodes.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary producer entries",
            source,
        })?;
    let mut total_bytes = 0usize;
    for mut node in producer_nodes {
        let mut existing = HashSet::new();
        existing
            .try_reserve(node.attrs.len())
            .map_err(|source| Error::Allocation {
                resource: "glossary producer MCE attribute index",
                source,
            })?;
        existing.extend(
            node.attrs
                .iter()
                .filter(|attribute| attribute.ns.as_ref() == MC)
                .map(|attribute| attribute.l.as_str()),
        );
        let mut missing = Vec::new();
        missing
            .try_reserve(inherited_mce.len())
            .map_err(|source| Error::Allocation {
                resource: "glossary missing MCE attributes",
                source,
            })?;
        missing.extend(
            inherited_mce
                .iter()
                .filter(|attribute| !existing.contains(attribute.l.as_str())),
        );
        drop(existing);
        node.attrs
            .try_reserve(missing.len())
            .map_err(|source| Error::Allocation {
                resource: "glossary inherited MCE attributes",
                source,
            })?;
        for attribute in missing {
            node.attrs.push(attribute.clone());
        }
        let xml = String::from_utf8(node_xml(&node, conformance == Conformance::Strict)?)
            .map_err(xml_error)?;
        total_bytes = total_bytes
            .checked_add(xml.len())
            .ok_or_else(|| invalid("glossary producer snapshot size overflow"))?;
        if total_bytes > MAX {
            return Err(invalid(
                "glossary producer snapshots exceed the 32 MiB aggregate limit",
            ));
        }
        let refs = relationship_references(&node)?;
        snapshots.push(ProducerEntry {
            conformance,
            xml: Arc::from(xml),
            refs,
        });
    }
    Ok(Some(snapshots))
}

fn attach_producer_entries(catalog: &mut Catalog, producer_entries: Option<Vec<ProducerEntry>>) {
    let Some(producer_entries) = producer_entries else {
        return;
    };
    if producer_entries.len() != catalog.entries.len() {
        return;
    }
    for (entry, producer) in catalog.entries.iter_mut().zip(producer_entries) {
        entry.producer = Some(Box::new(producer));
    }
}

fn merge_mce_attributes(output: &mut Vec<Attr>, node: &Node) -> Result<()> {
    let incoming = node
        .attrs
        .iter()
        .filter(|attribute| attribute.ns.as_ref() == MC)
        .count();
    let capacity = output
        .len()
        .checked_add(incoming)
        .ok_or_else(|| invalid("glossary MCE attribute count overflow"))?;
    let mut positions = HashMap::new();
    positions
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "glossary MCE attribute index",
            source,
        })?;
    for (index, attribute) in output.iter().enumerate() {
        positions.insert(attribute.l.clone(), index);
    }
    output
        .try_reserve(incoming)
        .map_err(|source| Error::Allocation {
            resource: "glossary MCE attribute scope",
            source,
        })?;
    for attribute in node
        .attrs
        .iter()
        .filter(|attribute| attribute.ns.as_ref() == MC)
    {
        if let Some(index) = positions.get(attribute.l.as_str()).copied() {
            output[index] = attribute.clone();
        } else {
            positions.insert(attribute.l.clone(), output.len());
            output.push(attribute.clone());
        }
    }
    Ok(())
}

fn parse_entry(n: &Node, semantic_bytes: &mut usize) -> Result<Entry> {
    expect(n, "docPart")?;
    noattrs(n)?;
    let c = kids(n)?;
    if c.len() > 2 {
        return Err(invalid("docPart has too many children"));
    }
    let mut props = None;
    let mut body = None;
    let mut i = 0;
    if let Some(properties) = c.first().filter(|node| node.l == "docPartPr") {
        props = Some(parse_props(properties)?);
        i = 1;
    }
    if i < c.len() {
        let body_node = c
            .get(i)
            .copied()
            .ok_or_else(|| invalid("document-part child index is out of bounds"))?;
        expect(body_node, "docPartBody")?;
        let xml = node_xml(body_node, false)?;
        add_semantic_xml_bytes(semantic_bytes, xml.len())?;
        body = Some(xml);
        i += 1;
    }
    if i != c.len() {
        return Err(invalid("invalid docPart child order"));
    }
    Ok(Entry {
        props,
        body,
        producer: None,
        sizes: [0; 2],
        refs: Arc::from([]),
        lineage: None,
    })
}
fn parse_props(n: &Node) -> Result<Props> {
    expect(n, "docPartPr")?;
    noattrs(n)?;
    let mut name_node = None;
    let mut style_node = None;
    let mut category_node = None;
    let mut types_node = None;
    let mut behaviors_node = None;
    let mut description_node = None;
    let mut id_node = None;
    for c in kids(n)? {
        expect_w(c)?;
        match c.l.as_str() {
            "name" => name_node = Some(c),
            "style" => style_node = Some(c),
            "category" => category_node = Some(c),
            "types" => types_node = Some(c),
            "behaviors" => behaviors_node = Some(c),
            "description" => description_node = Some(c),
            "guid" => id_node = Some(c),
            _ => return Err(invalid("unexpected docPartPr child")),
        }
    }
    let name = name_node
        .map(|name_node| {
            let value = wattr_get(name_node, "val")?
                .ok_or_else(|| invalid("building-block name is missing w:val"))?;
            bounded(&value)?;
            let name = Name {
                key: Arc::from(name_key(&value)?),
                value,
                decorated: onoff(name_node, "decorated")?,
            };
            only_w(name_node, &["val", "decorated"])?;
            leaf(name_node)?;
            Ok::<Name, Error>(name)
        })
        .transpose()?;

    let style = style_node.map(wval).transpose()?;
    let category = category_node
        .map(|node| {
            noattrs(node)?;
            let children = kids(node)?;
            let [category_name, gallery] = children.as_slice() else {
                return Err(invalid("category requires name then gallery"));
            };
            expect(category_name, "name")?;
            expect(gallery, "gallery")?;
            Ok::<Category, Error>(Category {
                name: wval(category_name)?,
                gallery: Gallery::new(wval(gallery)?)?,
            })
        })
        .transpose()?;
    let (all_kinds, kinds) = if let Some(node) = types_node {
        let all = onoff(node, "all")?;
        only_w(node, &["all"])?;
        let mut values = Kind::empty();
        let children = kids(node)?;
        for (count, child) in children.into_iter().enumerate() {
            if count >= MAX_VALUES {
                return Err(invalid("type limit exceeded"));
            }
            expect(child, "type")?;
            let value = parse_type(&wval(child)?)?;
            values.insert(value);
        }
        (all, values)
    } else {
        (None, Kind::empty())
    };
    let inserts = if let Some(node) = behaviors_node {
        noattrs(node)?;
        let mut values = Insert::empty();
        let children = kids(node)?;
        if children.is_empty() {
            return Err(invalid(
                "document-part behaviors requires at least one behavior",
            ));
        }
        for (count, child) in children.into_iter().enumerate() {
            if count >= MAX_VALUES {
                return Err(invalid("behavior limit exceeded"));
            }
            expect(child, "behavior")?;
            let value = parse_behavior(&wval(child)?)?;
            values.insert(value);
        }
        values
    } else {
        Insert::empty()
    };
    let description = description_node.map(wval).transpose()?;
    let id = id_node
        .map(|node| {
            let value = wattr_get(node, "val")?;
            only_w(node, &["val"])?;
            leaf(node)?;
            value.map(Id::new).transpose()
        })
        .transpose()?
        .flatten();
    Ok(Props {
        name,
        style,
        category,
        all_kinds,
        kinds,
        inserts,
        description,
        id,
    })
}

struct WritePlan {
    background: Option<Node>,
    bodies: Vec<Option<Node>>,
    producer_entries: Vec<Option<Arc<str>>>,
    bytes: usize,
}

#[derive(Default)]
struct XmlSize {
    bytes: usize,
}

trait XmlSink {
    fn push_str(&mut self, value: &str) -> Result<()>;
    fn push_char(&mut self, value: char) -> Result<()>;
}

impl XmlSink for XmlSize {
    fn push_str(&mut self, value: &str) -> Result<()> {
        self.bytes = self
            .bytes
            .checked_add(value.len())
            .ok_or_else(|| invalid("glossary XML output size overflow"))?;
        if self.bytes > MAX {
            return Err(invalid("serialized glossary document exceeds 32 MiB"));
        }
        Ok(())
    }

    fn push_char(&mut self, value: char) -> Result<()> {
        self.bytes = self
            .bytes
            .checked_add(value.len_utf8())
            .ok_or_else(|| invalid("glossary XML output size overflow"))?;
        if self.bytes > MAX {
            return Err(invalid("serialized glossary document exceeds 32 MiB"));
        }
        Ok(())
    }
}

impl XmlSink for String {
    fn push_str(&mut self, value: &str) -> Result<()> {
        String::push_str(self, value);
        Ok(())
    }

    fn push_char(&mut self, value: char) -> Result<()> {
        String::push(self, value);
        Ok(())
    }
}

fn entry_analysis(entry: &Entry) -> Result<EntryAnalysis> {
    validate_entry_fields(entry)?;
    let body = entry
        .body
        .as_deref()
        .map(|xml| prepare_opaque(xml, "docPartBody"))
        .transpose()?;
    let refs = body
        .as_ref()
        .map(relationship_references)
        .transpose()?
        .unwrap_or_else(|| Arc::from([]));
    let mut sizes = [0usize; 2];
    for conformance in [Conformance::Transitional, Conformance::Strict] {
        let mut size = XmlSize::default();
        if let Some(producer) = entry
            .producer
            .as_deref()
            .filter(|producer| producer.conformance == conformance)
        {
            size.push_str(&producer.xml)?;
        } else {
            write_entry(&mut size, entry, body.as_ref(), conformance)?;
        }
        sizes[conformance.index()] = size.bytes;
    }
    Ok(EntryAnalysis { sizes, refs })
}

fn background_analysis(background: Option<&[u8]>) -> Result<EntryAnalysis> {
    let mut sizes = [0usize; 2];
    let Some(background) = background else {
        return Ok(EntryAnalysis {
            sizes,
            refs: Arc::from([]),
        });
    };
    let node = prepare_opaque(background, "background")?;
    let refs = relationship_references(&node)?;
    for conformance in [Conformance::Transitional, Conformance::Strict] {
        let mut size = XmlSize::default();
        node_write(&mut size, &node, conformance == Conformance::Strict)?;
        sizes[conformance.index()] = size.bytes;
    }
    Ok(EntryAnalysis { sizes, refs })
}

fn add_sizes(left: [usize; 2], right: [usize; 2]) -> Result<[usize; 2]> {
    Ok([
        left[0]
            .checked_add(right[0])
            .ok_or_else(|| invalid("Transitional glossary size overflow"))?,
        left[1]
            .checked_add(right[1])
            .ok_or_else(|| invalid("Strict glossary size overflow"))?,
    ])
}

fn replace_sizes(total: [usize; 2], old: [usize; 2], new: [usize; 2]) -> Result<[usize; 2]> {
    Ok([
        total[0]
            .checked_sub(old[0])
            .and_then(|value| value.checked_add(new[0]))
            .ok_or_else(|| invalid("Transitional glossary replacement size overflow"))?,
        total[1]
            .checked_sub(old[1])
            .and_then(|value| value.checked_add(new[1]))
            .ok_or_else(|| invalid("Strict glossary replacement size overflow"))?,
    ])
}

fn validate_catalog_sizes(
    entry_bytes: [usize; 2],
    background_bytes: [usize; 2],
    entry_count: usize,
) -> Result<()> {
    for conformance in [Conformance::Transitional, Conformance::Strict] {
        let mut size = XmlSize::default();
        write_catalog_open(&mut size, conformance)?;
        if entry_count != 0 {
            size.push_str("<w:docParts></w:docParts>")?;
        }
        size.push_str("</w:glossaryDocument>")?;
        let index = conformance.index();
        let bytes = size
            .bytes
            .checked_add(background_bytes[index])
            .and_then(|bytes| bytes.checked_add(entry_bytes[index]))
            .ok_or_else(|| invalid("glossary catalog size overflow"))?;
        if bytes > MAX {
            return Err(invalid("serialized glossary document exceeds 32 MiB"));
        }
    }
    Ok(())
}

fn write_catalog_open<X: XmlSink>(x: &mut X, conformance: Conformance) -> Result<()> {
    x.push_str(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:glossaryDocument xmlns:w=""#,
    )?;
    x.push_str(conformance.word())?;
    x.push_str(r#"" xmlns:r=""#)?;
    x.push_str(conformance.relationships())?;
    x.push_str(r#"">"#)
}

fn plan_write(value: &Catalog, conformance: Conformance) -> Result<WritePlan> {
    validate_catalog_fields(value)?;
    let background = value
        .background
        .as_deref()
        .map(|xml| prepare_opaque_for(xml, "background", conformance))
        .transpose()?;
    let mut bodies = Vec::new();
    bodies
        .try_reserve_exact(value.entries.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary XML write plan",
            source,
        })?;
    let mut producer_entries = Vec::new();
    producer_entries
        .try_reserve_exact(value.entries.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary producer write plan",
            source,
        })?;
    for entry in &value.entries {
        let producer = entry
            .producer
            .as_deref()
            .filter(|producer| producer.conformance == conformance);
        producer_entries.push(producer.map(|producer| Arc::clone(&producer.xml)));
        bodies.push(if producer.is_some() {
            None
        } else {
            entry
                .body
                .as_deref()
                .map(|xml| prepare_opaque_for(xml, "docPartBody", conformance))
                .transpose()?
        });
    }
    let mut plan = WritePlan {
        background,
        bodies,
        producer_entries,
        bytes: 0,
    };
    let mut size = XmlSize::default();
    emit_catalog(&mut size, value, conformance, &plan)?;
    plan.bytes = size.bytes;
    Ok(plan)
}

fn emit_catalog<X: XmlSink>(
    x: &mut X,
    value: &Catalog,
    conformance: Conformance,
    plan: &WritePlan,
) -> Result<()> {
    write_catalog_open(x, conformance)?;
    match (&value.background, &plan.background) {
        (Some(_), Some(background)) => {
            node_write(x, background, conformance == Conformance::Strict)?;
        },
        (None, None) => {},
        _ => return Err(invalid("glossary background write plan is inconsistent")),
    }
    let mut bodies = plan.bodies.iter();
    let mut producer_entries = plan.producer_entries.iter();
    if !value.entries.is_empty() {
        x.push_str("<w:docParts>")?;
        for entry in &value.entries {
            let body = bodies
                .next()
                .ok_or_else(|| invalid("glossary entry write plan is incomplete"))?;
            let producer = producer_entries
                .next()
                .ok_or_else(|| invalid("glossary producer write plan is incomplete"))?;
            if let Some(producer) = producer {
                if body.is_some() {
                    return Err(invalid("glossary producer write plan has a duplicate body"));
                }
                x.push_str(producer)?;
            } else {
                write_entry(x, entry, body.as_ref(), conformance)?;
            }
        }
        x.push_str("</w:docParts>")?;
    }
    if bodies.next().is_some() {
        return Err(invalid("glossary entry write plan has unused bodies"));
    }
    if producer_entries.next().is_some() {
        return Err(invalid("glossary producer write plan has unused entries"));
    }
    x.push_str("</w:glossaryDocument>")?;
    Ok(())
}

fn write_entry<X: XmlSink>(
    x: &mut X,
    e: &Entry,
    body: Option<&Node>,
    conformance: Conformance,
) -> Result<()> {
    x.push_str("<w:docPart>")?;
    if let Some(p) = &e.props {
        write_props(x, p)?;
    }
    match (&e.body, body) {
        (Some(_), Some(body)) => node_write(x, body, conformance == Conformance::Strict)?,
        (None, None) => {},
        _ => return Err(invalid("building-block body write plan is inconsistent")),
    }
    x.push_str("</w:docPart>")?;
    Ok(())
}
fn write_props<X: XmlSink>(x: &mut X, p: &Props) -> Result<()> {
    x.push_str("<w:docPartPr>")?;
    if let Some(name) = &p.name {
        x.push_str("<w:name")?;
        wattr(x, "val", name.as_str())?;
        wonoff(x, "decorated", name.decorated())?;
        x.push_str("/>")?;
    }
    if let Some(v) = &p.style {
        leafval(x, "style", v)?;
    }
    if let Some(c) = &p.category {
        x.push_str("<w:category>")?;
        leafval(x, "name", &c.name)?;
        leafval(x, "gallery", c.gallery.as_str())?;
        x.push_str("</w:category>")?;
    }
    if p.all_kinds.is_some() || !p.kinds.is_empty() {
        x.push_str("<w:types")?;
        wonoff(x, "all", p.all_kinds)?;
        x.push_char('>')?;
        for (kind, value) in KIND_VALUES {
            if p.kinds.contains(kind) {
                leafval(x, "type", value)?;
            }
        }
        x.push_str("</w:types>")?;
    }
    if !p.inserts.is_empty() {
        x.push_str("<w:behaviors>")?;
        for (insert, value) in INSERT_VALUES {
            if p.inserts.contains(insert) {
                leafval(x, "behavior", value)?;
            }
        }
        x.push_str("</w:behaviors>")?;
    }
    if let Some(v) = &p.description {
        leafval(x, "description", v)?;
    }
    if let Some(v) = &p.id {
        leafval(x, "guid", v.as_str())?;
    }
    x.push_str("</w:docPartPr>")?;
    Ok(())
}
fn leafval<X: XmlSink>(x: &mut X, n: &str, v: &str) -> Result<()> {
    x.push_str("<w:")?;
    x.push_str(n)?;
    wattr(x, "val", v)?;
    x.push_str("/>")
}
fn wattr<X: XmlSink>(x: &mut X, n: &str, v: &str) -> Result<()> {
    x.push_str(" w:")?;
    x.push_str(n)?;
    x.push_str("=\"")?;
    esc(x, v)?;
    x.push_char('"')
}
fn wonoff<X: XmlSink>(x: &mut X, n: &str, v: Option<bool>) -> Result<()> {
    if let Some(v) = v {
        wattr(x, n, if v { "1" } else { "0" })?;
    }
    Ok(())
}

fn validate_catalog_fields(v: &Catalog) -> Result<()> {
    if v.entries.len() > MAX_PARTS {
        return Err(invalid("glossary entry limit exceeded"));
    }
    for e in &v.entries {
        validate_entry_fields(e)?;
    }
    Ok(())
}

fn validate_entry_fields(entry: &Entry) -> Result<()> {
    let Some(props) = &entry.props else {
        return Ok(());
    };
    if let Some(name) = &props.name {
        bounded(name.as_str())?;
    }
    for value in [
        props.style.as_deref(),
        props.category.as_ref().map(Category::name),
        props.description.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        bounded(value)?;
    }
    if !Kind::all().contains(props.kinds) || !Insert::all().contains(props.inserts) {
        return Err(invalid("unknown glossary option flag"));
    }
    if props.all_kinds.is_some() && props.kinds.is_empty() {
        return Err(invalid(
            "document-part types requires at least one kind when present",
        ));
    }
    Ok(())
}

fn validate_authored_name(entry: &Entry) -> Result<()> {
    let Some(name) = entry.props.as_ref().and_then(|props| props.name.as_ref()) else {
        return Err(invalid(
            "authored building block requires properties and a name",
        ));
    };
    validate_name(name.as_str())
}

fn authored_analysis(entry: &Entry) -> Result<EntryAnalysis> {
    validate_authored_name(entry)?;
    entry_analysis(entry)
}

fn prepare_opaque(xml: &[u8], local: &str) -> Result<Node> {
    if xml.len() > MAX {
        return Err(invalid(format!("glossary {local} payload exceeds 32 MiB")));
    }
    let root = parse_dom(xml)?;
    expect(&root, local)?;
    Ok(root)
}

fn prepare_opaque_for(xml: &[u8], local: &str, conformance: Conformance) -> Result<Node> {
    let root = prepare_opaque(xml, local)?;
    if conformance == Conformance::Strict {
        reject_vml(&root)?;
    }
    Ok(root)
}

fn reject_vml(node: &Node) -> Result<()> {
    if node.ns.as_ref() == VML
        || node
            .attrs
            .iter()
            .any(|attribute| attribute.ns.as_ref() == VML)
    {
        return Err(invalid("Strict glossary content cannot contain VML"));
    }
    for content in &node.content {
        if let Content::Node(child) = content {
            reject_vml(child)?;
        }
    }
    Ok(())
}

fn validate_name(value: &str) -> Result<()> {
    bounded(value)?;
    if value.trim().is_empty() {
        Err(invalid("building-block name cannot be empty"))
    } else {
        Ok(())
    }
}

fn validate_raw_part(name: &str, content_type: &str, len: usize) -> Result<()> {
    let uri = validate_physical_part(name, content_type, len)?;
    if uri.as_str() == "/word/glossary/document.xml" {
        return Err(invalid(
            "glossary auxiliary part conflicts with the default root",
        ));
    }
    Ok(())
}

fn validate_physical_part(name: &str, content_type: &str, len: usize) -> Result<PackURI> {
    let uri = PackURI::new(name).map_err(Error::Uri)?;
    if is_signature_part(&uri) || is_reserved_physical_part(&uri) {
        return Err(invalid(format!(
            "'{}' is reserved OPC package infrastructure",
            uri.as_str()
        )));
    }
    ContentType::new(content_type.to_owned())?;
    let media_type = content_type.split(';').next().unwrap_or_default().trim();
    if [
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN,
        ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE,
        ct::OPC_DIGITAL_SIGNATURE_CERTIFICATE,
        ct::OPC_RELATIONSHIPS,
    ]
    .iter()
    .any(|reserved| media_type.eq_ignore_ascii_case(reserved))
    {
        return Err(invalid(
            "glossary parts cannot use reserved OPC infrastructure content types",
        ));
    }
    if len > MAX_GRAPH_BYTES {
        return Err(invalid("glossary auxiliary part exceeds 256 MiB"));
    }
    Ok(uri)
}

fn name_key(value: &str) -> Result<String> {
    bounded(value)?;
    let mut key = String::new();
    key.try_reserve(value.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary semantic-name key",
            source,
        })?;
    for character in value.chars().nfd().default_case_fold().nfd() {
        let len = key
            .len()
            .checked_add(character.len_utf8())
            .ok_or_else(|| invalid("glossary semantic-name key size overflow"))?;
        if len > MAX_NAME_KEY {
            return Err(invalid(
                "glossary semantic-name key exceeds the normalized size limit",
            ));
        }
        if key.capacity() < len {
            key.try_reserve(character.len_utf8())
                .map_err(|source| Error::Allocation {
                    resource: "glossary semantic-name key",
                    source,
                })?;
        }
        key.push(character);
    }
    Ok(key)
}

fn canonical_id(value: &str) -> bool {
    let bytes = value.as_bytes();
    bytes.len() == 38
        && bytes.first() == Some(&b'{')
        && bytes.last() == Some(&b'}')
        && bytes.iter().enumerate().all(|(index, byte)| match index {
            0 => *byte == b'{',
            9 | 14 | 19 | 24 => *byte == b'-',
            37 => *byte == b'}',
            _ => byte.is_ascii_digit() || (b'A'..=b'F').contains(byte),
        })
}

fn valid_ncname(value: &str) -> bool {
    let mut characters = value.chars();
    characters.next().is_some_and(is_ncname_start) && characters.all(is_ncname_char)
}

fn is_ncname_start(character: char) -> bool {
    matches!(
        character,
        'A'..='Z'
            | '_'
            | 'a'..='z'
            | '\u{00c0}'..='\u{00d6}'
            | '\u{00d8}'..='\u{00f6}'
            | '\u{00f8}'..='\u{02ff}'
            | '\u{0370}'..='\u{037d}'
            | '\u{037f}'..='\u{1fff}'
            | '\u{200c}'..='\u{200d}'
            | '\u{2070}'..='\u{218f}'
            | '\u{2c00}'..='\u{2fef}'
            | '\u{3001}'..='\u{d7ff}'
            | '\u{f900}'..='\u{fdcf}'
            | '\u{fdf0}'..='\u{fffd}'
            | '\u{10000}'..='\u{effff}'
    )
}

fn is_ncname_char(character: char) -> bool {
    is_ncname_start(character)
        || matches!(
            character,
            '-' | '.' | '0'..='9' | '\u{00b7}' | '\u{0300}'..='\u{036f}' | '\u{203f}'..='\u{2040}'
        )
}

fn node_xml(n: &Node, s: bool) -> Result<Vec<u8>> {
    let mut size = XmlSize::default();
    node_write(&mut size, n, s)?;
    let mut x = String::new();
    x.try_reserve_exact(size.bytes)
        .map_err(|source| Error::Allocation {
            resource: "glossary opaque XML",
            source,
        })?;
    node_write(&mut x, n, s)?;
    if x.len() != size.bytes {
        return Err(invalid("glossary opaque XML plan did not match output"));
    }
    Ok(x.into_bytes())
}
fn node_write<X: XmlSink>(x: &mut X, n: &Node, s: bool) -> Result<()> {
    node_write_inner(x, n, s, true)
}

fn node_write_inner<X: XmlSink>(x: &mut X, n: &Node, s: bool, root: bool) -> Result<()> {
    x.push_char('<')?;
    x.push_str(&n.q)?;
    if root {
        for (prefix, namespace) in namespace_scope(&n.bindings)? {
            write_namespace_binding(x, prefix, namespace, s)?;
        }
    } else {
        for (prefix, namespace) in &n.bindings.local {
            write_namespace_binding(x, prefix, namespace, s)?;
        }
    }
    for a in &n.attrs {
        x.push_char(' ')?;
        x.push_str(&a.q)?;
        x.push_str("=\"")?;
        esc(x, &a.v)?;
        x.push_char('"')?;
    }
    if n.content.is_empty() {
        x.push_str("/>")?;
        return Ok(());
    }
    x.push_char('>')?;
    for c in &n.content {
        match c {
            Content::Node(n) => node_write_inner(x, n, s, false)?,
            Content::Text(v) => text(x, v)?,
            Content::CData(v) => {
                if v.contains('\r') || v.contains("]]>") {
                    text(x, v)?;
                } else {
                    x.push_str("<![CDATA[")?;
                    x.push_str(v)?;
                    x.push_str("]]>")?;
                }
            },
            Content::Comment(v) => {
                x.push_str("<!--")?;
                x.push_str(v)?;
                x.push_str("-->")?;
            },
        }
    }
    x.push_str("</")?;
    x.push_str(&n.q)?;
    x.push_char('>')?;
    Ok(())
}

fn namespace_scope(bindings: &NamespaceFrame) -> Result<Vec<(&str, &str)>> {
    let mut frames = Vec::new();
    frames
        .try_reserve_exact(MAX_DEPTH)
        .map_err(|source| Error::Allocation {
            resource: "glossary namespace frames",
            source,
        })?;
    let mut current = Some(bindings);
    while let Some(frame) = current {
        frames.push(frame);
        current = frame.parent.as_deref();
    }
    frames.reverse();

    let mut result = Vec::new();
    result
        .try_reserve(bindings.declaration_count)
        .map_err(|source| Error::Allocation {
            resource: "glossary namespace scope",
            source,
        })?;
    let mut positions = HashMap::new();
    positions
        .try_reserve(bindings.declaration_count)
        .map_err(|source| Error::Allocation {
            resource: "glossary namespace scope index",
            source,
        })?;
    for frame in frames {
        for (prefix, namespace) in &frame.local {
            if let Some(index) = positions.get(prefix.as_str()).copied() {
                result[index] = (prefix.as_str(), namespace.as_ref());
            } else {
                positions.insert(prefix.as_str(), result.len());
                result.push((prefix.as_str(), namespace.as_ref()));
            }
        }
    }
    Ok(result)
}

fn write_namespace_binding<X: XmlSink>(
    x: &mut X,
    prefix: &str,
    namespace: &str,
    strict: bool,
) -> Result<()> {
    if prefix.is_empty() {
        x.push_str(" xmlns=\"")?
    } else {
        x.push_str(" xmlns:")?;
        x.push_str(prefix)?;
        x.push_str("=\"")?
    }
    esc(x, mapns(namespace, strict))?;
    x.push_char('"')
}
fn mapns(v: &str, s: bool) -> &str {
    if s {
        match v {
            W => WS,
            R => RS,
            _ => v,
        }
    } else {
        match v {
            WS => W,
            RS => R,
            _ => v,
        }
    }
}

fn kids(n: &Node) -> Result<Vec<&Node>> {
    let mut v = Vec::new();
    for c in &n.content {
        match c {
            Content::Node(x) => v.push(x),
            Content::Text(x) if x.trim().is_empty() => {},
            Content::Comment(_) => {},
            _ => return Err(invalid("unexpected text in typed glossary metadata")),
        }
    }
    Ok(v)
}
fn leaf(n: &Node) -> Result<()> {
    if kids(n)?.is_empty() {
        Ok(())
    } else {
        Err(invalid("glossary metadata leaf has children"))
    }
}
fn validate_word_dialect(n: &Node, conformance: Conformance) -> Result<()> {
    if matches!(n.ns.as_ref(), W | WS) && n.ns.as_ref() != conformance.word() {
        return Err(invalid(
            "glossary mixes Strict and Transitional WordprocessingML",
        ));
    }
    if matches!(n.ns.as_ref(), R | RS) && n.ns.as_ref() != conformance.relationships() {
        return Err(invalid(
            "glossary mixes Strict and Transitional relationship namespaces",
        ));
    }
    if conformance == Conformance::Strict && n.ns.as_ref() == VML {
        return Err(invalid("Strict glossary content cannot contain VML"));
    }
    for attribute in &n.attrs {
        if matches!(attribute.ns.as_ref(), W | WS) && attribute.ns.as_ref() != conformance.word() {
            return Err(invalid(
                "glossary mixes Strict and Transitional WordprocessingML attributes",
            ));
        }
        if matches!(attribute.ns.as_ref(), R | RS)
            && attribute.ns.as_ref() != conformance.relationships()
        {
            return Err(invalid(
                "glossary mixes Strict and Transitional relationship attributes",
            ));
        }
        if conformance == Conformance::Strict && attribute.ns.as_ref() == VML {
            return Err(invalid(
                "Strict glossary content cannot contain VML attributes",
            ));
        }
    }
    for content in &n.content {
        if let Content::Node(child) = content {
            validate_word_dialect(child, conformance)?;
        }
    }
    Ok(())
}
fn expect(n: &Node, l: &str) -> Result<()> {
    if matches!(n.ns.as_ref(), W | WS) && n.l == l {
        Ok(())
    } else {
        Err(invalid(format!("expected WordprocessingML {l}")))
    }
}
fn expect_w(n: &Node) -> Result<()> {
    if matches!(n.ns.as_ref(), W | WS) {
        Ok(())
    } else {
        Err(invalid("expected WordprocessingML metadata"))
    }
}
fn wval(n: &Node) -> Result<String> {
    let v = wattr_get(n, "val")?.ok_or_else(|| invalid("missing w:val"))?;
    bounded(&v)?;
    only_w(n, &["val"])?;
    leaf(n)?;
    Ok(v)
}
fn wattr_get(n: &Node, l: &str) -> Result<Option<String>> {
    let mut v = None;
    for a in &n.attrs {
        if matches!(a.ns.as_ref(), W | WS) && a.l == l {
            if v.is_some() {
                return Err(invalid("duplicate WordprocessingML attribute"));
            }
            v = Some(a.v.clone());
        }
    }
    Ok(v)
}
fn onoff(n: &Node, l: &str) -> Result<Option<bool>> {
    let value = wattr_get(n, l)?;
    match value.as_deref() {
        None => Ok(None),
        Some("1" | "true") => Ok(Some(true)),
        Some("0" | "false") => Ok(Some(false)),
        Some("on") if n.ns.as_ref() == W => Ok(Some(true)),
        Some("off") if n.ns.as_ref() == W => Ok(Some(false)),
        _ => Err(invalid(format!("invalid on/off attribute '{l}'"))),
    }
}
fn only_w(n: &Node, allowed: &[&str]) -> Result<()> {
    for a in &n.attrs {
        if !(matches!(a.ns.as_ref(), W | WS) && allowed.contains(&a.l.as_str())) {
            return Err(invalid(format!("unexpected glossary attribute '{}'", a.q)));
        }
    }
    Ok(())
}
fn noattrs(n: &Node) -> Result<()> {
    if n.attrs.is_empty() {
        Ok(())
    } else {
        Err(invalid("unexpected glossary attributes"))
    }
}
fn parse_type(v: &str) -> Result<Kind> {
    match v {
        "none" => Ok(Kind::NONE),
        "normal" => Ok(Kind::NORMAL),
        "autoExp" => Ok(Kind::AUTO_EXPAND),
        "toolbar" => Ok(Kind::TOOLBAR),
        "speller" => Ok(Kind::SPELLER),
        "formFld" => Ok(Kind::FORM_FIELD),
        "bbPlcHdr" => Ok(Kind::SDT_PLACEHOLDER),
        _ => Err(invalid(format!("invalid document-part type '{v}'"))),
    }
}
const KIND_VALUES: [(Kind, &str); 7] = [
    (Kind::NONE, "none"),
    (Kind::NORMAL, "normal"),
    (Kind::AUTO_EXPAND, "autoExp"),
    (Kind::TOOLBAR, "toolbar"),
    (Kind::SPELLER, "speller"),
    (Kind::FORM_FIELD, "formFld"),
    (Kind::SDT_PLACEHOLDER, "bbPlcHdr"),
];
fn parse_behavior(v: &str) -> Result<Insert> {
    match v {
        "content" => Ok(Insert::CONTENT),
        "p" => Ok(Insert::PARAGRAPH),
        "pg" => Ok(Insert::PAGE),
        _ => Err(invalid(format!("invalid insertion behavior '{v}'"))),
    }
}
const INSERT_VALUES: [(Insert, &str); 3] = [
    (Insert::CONTENT, "content"),
    (Insert::PARAGRAPH, "p"),
    (Insert::PAGE, "pg"),
];
const GALLERIES: &[&str] = &[
    "placeholder",
    "any",
    "default",
    "docParts",
    "coverPg",
    "eq",
    "ftrs",
    "hdrs",
    "pgNum",
    "tbls",
    "watermarks",
    "autoTxt",
    "txtBox",
    "pgNumT",
    "pgNumB",
    "pgNumMargins",
    "tblOfContents",
    "bib",
    "custQuickParts",
    "custCoverPg",
    "custEq",
    "custFtrs",
    "custHdrs",
    "custPgNum",
    "custTbls",
    "custWatermarks",
    "custAutoTxt",
    "custTxtBox",
    "custPgNumT",
    "custPgNumB",
    "custPgNumMargins",
    "custTblOfContents",
    "custBib",
    "custom1",
    "custom2",
    "custom3",
    "custom4",
    "custom5",
];
fn split(q: &str) -> Result<(&str, &str)> {
    if let Some((p, l)) = q.split_once(':') {
        if l.is_empty() || l.contains(':') {
            return Err(invalid("invalid QName"));
        }
        Ok((p, l))
    } else {
        Ok(("", q))
    }
}
fn bounded(v: &str) -> Result<()> {
    if v.len() > MAX_STRING {
        return Err(invalid("glossary metadata string exceeds 1 MiB"));
    }
    if !v.chars().all(xml_char) {
        return Err(invalid(
            "glossary metadata contains a character forbidden by XML 1.0",
        ));
    }
    Ok(())
}

fn xml_char(character: char) -> bool {
    matches!(
        character,
        '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
    )
}
fn esc<X: XmlSink>(x: &mut X, v: &str) -> Result<()> {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;")?,
            '<' => x.push_str("&lt;")?,
            '"' => x.push_str("&quot;")?,
            '\r' => x.push_str("&#xD;")?,
            '\n' => x.push_str("&#xA;")?,
            '\t' => x.push_str("&#x9;")?,
            _ => x.push_char(c)?,
        }
    }
    Ok(())
}
fn text<X: XmlSink>(x: &mut X, v: &str) -> Result<()> {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;")?,
            '<' => x.push_str("&lt;")?,
            '>' => x.push_str("&gt;")?,
            '\r' => x.push_str("&#xD;")?,
            _ => x.push_char(c)?,
        }
    }
    Ok(())
}
fn invalid(value: impl Into<String>) -> Error {
    Error::Invalid(value.into())
}
fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use std::sync::Arc;

    fn fixture(bytes: &[u8]) -> (Catalog, Conformance) {
        let package = OpcPackage::from_bytes(bytes).expect("package");
        load(&package).expect("load").expect("glossary")
    }

    fn package_for(conformance: Conformance) -> OpcPackage {
        let mut package = OpcPackage::new();
        let (word, office_document) = match conformance {
            Conformance::Transitional => (W, rt::OFFICE_DOCUMENT),
            Conformance::Strict => (WS, rt::STRICT_OFFICE_DOCUMENT),
        };
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/word/document.xml").expect("URI"),
                ct::WML_DOCUMENT_MAIN.to_owned(),
                format!(r#"<w:document xmlns:w="{word}"><w:body/></w:document>"#).into_bytes(),
            )))
            .expect("main part");
        package.relate_to("word/document.xml", office_document);
        package
    }

    fn package() -> OpcPackage {
        package_for(Conformance::Transitional)
    }

    fn add_empty_glossary(package: &mut OpcPackage, conformance: Conformance) -> PackURI {
        let root = PackURI::new("/word/glossary/document.xml").unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                root.clone(),
                CT.to_owned(),
                write(&Catalog::new(), conformance).unwrap(),
            )))
            .unwrap();
        let main = package.main_document_part().unwrap().partname().clone();
        package
            .get_part_mut(&main)
            .unwrap()
            .rels_mut()
            .get_or_add(conformance.glossary_relationship(), "glossary/document.xml");
        root
    }

    fn mark_signed(package: &mut OpcPackage) {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                Vec::new(),
            )))
            .unwrap();
        package.rels_mut().add_relationship(
            rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rSignature".to_owned(),
            false,
        );
    }

    fn entry(name: &str) -> Entry {
        Entry::new(
            name,
            format!(r#"<w:docPartBody xmlns:w="{W}"><w:p/></w:docPartBody>"#).into_bytes(),
        )
        .expect("entry")
    }

    #[test]
    fn poi_placeholders_and_strict_roundtrip() {
        let (catalog, _) = fixture(include_bytes!(
            "../../../test-data/poi/test-data/document/Bug54849.docx"
        ));
        assert_eq!(catalog.len(), 3);
        let props = catalog.at(0).unwrap().props().unwrap();
        assert_eq!(props.kinds, Kind::SDT_PLACEHOLDER);
        assert_eq!(
            props.category.as_ref().unwrap().gallery().as_str(),
            "placeholder"
        );
        let xml = write(&catalog, Conformance::Strict).unwrap();
        assert_eq!(read(&xml).unwrap().0.len(), 3);
    }

    #[test]
    fn libreoffice_multiple_autotext_and_tables_stay_inert() {
        let (catalog, _) = fixture(include_bytes!(
            "../../../test-data/libreoffice-core/sw/qa/extras/uiwriter/data/autotext-multiple.dotx"
        ));
        assert_eq!(catalog.len(), 3);
        assert_eq!(catalog.at(0).unwrap().name(), Some("Multiple"));
        let body = std::str::from_utf8(catalog.at(2).unwrap().body().unwrap()).unwrap();
        assert!(body.contains("w:tbl"));
        assert!(body.contains("jksdjkfdskjfds"));
    }

    #[test]
    fn libreoffice_empty_glossary_is_valid() {
        let (catalog, conformance) = fixture(include_bytes!(
            "../../../test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/testGlossary.docx"
        ));
        assert!(catalog.is_empty());
        assert!(write(&catalog, conformance).is_ok());
    }

    #[test]
    fn xsd_all_is_order_independent_and_last_duplicate_wins() {
        let xml = format!(
            r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:description w:val="first"/><w:name w:val="First"/><w:types><w:type w:val="normal"/></w:types><w:name w:val="Last"/><w:description w:val="last"/></w:docPartPr><w:docPartBody><w:p/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        let (catalog, _) = read(xml.as_bytes()).unwrap();
        let entry = catalog.at(0).unwrap();
        assert_eq!(entry.name(), Some("Last"));
        assert_eq!(entry.props().unwrap().description.as_deref(), Some("last"));
        assert_eq!(entry.props().unwrap().kinds, Kind::NORMAL);
    }

    #[test]
    fn compact_option_flags_collapse_duplicates_and_reject_unknown_or_empty_values() {
        let xml = format!(
            r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Flags"/><w:types><w:type w:val="normal"/><w:type w:val="toolbar"/></w:types><w:behaviors><w:behavior w:val="content"/><w:behavior w:val="pg"/></w:behaviors></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        let (catalog, conformance) = read(xml.as_bytes()).unwrap();
        let props = catalog.at(0).unwrap().props().unwrap();
        assert_eq!(props.kinds, Kind::NORMAL | Kind::TOOLBAR);
        assert_eq!(props.inserts, Insert::CONTENT | Insert::PAGE);
        assert_eq!(
            read(&write(&catalog, conformance).unwrap()).unwrap().0,
            catalog
        );

        let duplicate = xml.replace(
            r#"<w:type w:val="toolbar"/>"#,
            r#"<w:type w:val="normal"/>"#,
        );
        let duplicate_catalog = read(duplicate.as_bytes()).unwrap().0;
        assert_eq!(
            duplicate_catalog.at(0).unwrap().props().unwrap().kinds,
            Kind::NORMAL
        );

        let word_empty_types = xml.replace(
            r#"<w:types><w:type w:val="normal"/><w:type w:val="toolbar"/></w:types>"#,
            "<w:types/>",
        );
        let (word_catalog, word_conformance) = read(word_empty_types.as_bytes()).unwrap();
        assert!(
            word_catalog
                .at(0)
                .unwrap()
                .props()
                .unwrap()
                .kinds
                .is_empty()
        );
        assert!(write(&word_catalog, word_conformance).is_ok());

        let empty_behaviors = xml.replace(
            r#"<w:behavior w:val="content"/><w:behavior w:val="pg"/>"#,
            "",
        );
        assert!(read(empty_behaviors.as_bytes()).is_err());

        let mut props = Props::new(Name::new("Unknown flag").unwrap());
        props.kinds = Kind::from_bits_retain(0x80);
        assert!(entry("Unknown flag").with_props(props).is_err());
    }

    #[test]
    fn xml_forbidden_metadata_is_rejected_without_catalog_mutation() {
        assert!(Name::new("bad\0name").is_err());
        assert!(Name::new("bad\u{fffe}name").is_err());
        assert!(Category::new("bad\0category", Gallery::new("autoTxt").unwrap()).is_err());

        let mut catalog = Catalog::new();
        catalog.add(entry("Keep")).unwrap();
        let before = catalog.clone();
        let mut invalid = Props::new(Name::new("Invalid").unwrap());
        invalid.description = Some("bad\0description".to_owned());
        assert!(entry("Invalid").with_props(invalid).is_err());
        assert_eq!(catalog, before);
    }

    #[test]
    fn invalid_earlier_duplicates_are_ignored_before_last_value_validation() {
        let xml = format!(
            r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:name/><w:guid w:val="invalid"/><w:types><w:type w:val="invalid"/></w:types><w:name w:val="Last"/><w:guid w:val="{{12345678-1234-4ABC-8DEF-1234567890AB}}"/><w:types><w:type w:val="normal"/></w:types></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        let (catalog, _) = read(xml.as_bytes()).unwrap();
        let props = catalog.at(0).unwrap().props().unwrap();
        assert_eq!(props.name().unwrap().as_str(), "Last");
        assert_eq!(
            props.id.as_ref().unwrap().as_str(),
            "{12345678-1234-4ABC-8DEF-1234567890AB}"
        );
        assert_eq!(props.kinds, Kind::NORMAL);
    }

    #[test]
    fn mixed_dialects_and_foreign_typed_qnames_are_rejected() {
        let mixed = format!(
            r#"<w:glossaryDocument xmlns:w="{W}" xmlns:s="{WS}"><w:docParts><w:docPart><w:docPartPr><s:name s:val="Mixed"/></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        assert!(read(mixed.as_bytes()).is_err());

        let spoofed = format!(
            r#"<w:glossaryDocument xmlns:w="{W}" xmlns:u="urn:spoof"><w:docParts><w:docPart><u:docPartPr><w:name w:val="Spoof"/></u:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        assert!(read(spoofed.as_bytes()).is_err());
    }

    #[test]
    fn duplicate_producer_names_are_readable_but_semantically_ambiguous() {
        let xml = format!(
            r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Résumé"/></w:docPartPr></w:docPart><w:docPart><w:docPartPr><w:name w:val="RÉSUMÉ"/></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        let (catalog, _) = read(xml.as_bytes()).unwrap();
        assert!(catalog.get("résumé").is_err());
        assert_eq!(catalog.at(0).unwrap().name(), Some("Résumé"));
        assert_eq!(catalog.at(1).unwrap().name(), Some("RÉSUMÉ"));
    }

    #[test]
    fn selectors_support_atomic_rename_replacement_and_numeric_ambiguity_repair() {
        let mut catalog = Catalog::new();
        catalog.add(entry("Alpha")).unwrap();
        catalog.add(entry("Beta")).unwrap();
        assert_eq!(catalog.get("ALPHA").unwrap().unwrap().name(), Some("Alpha"));

        let previous = catalog.replace("alpha", entry("Gamma")).unwrap().unwrap();
        assert_eq!(previous.name(), Some("Alpha"));
        assert!(catalog.get("alpha").unwrap().is_none());
        let body_pointer = catalog
            .get("gamma")
            .unwrap()
            .unwrap()
            .body()
            .unwrap()
            .as_ptr();
        assert!(
            catalog
                .rename("GAMMA", Name::new("Epsilon").unwrap())
                .unwrap()
        );
        assert_eq!(
            catalog
                .get("epsilon")
                .unwrap()
                .unwrap()
                .body()
                .unwrap()
                .as_ptr(),
            body_pointer
        );
        let before = catalog.clone();
        assert!(
            catalog
                .rename("epsilon", Name::new("Beta").unwrap())
                .is_err()
        );
        assert_eq!(catalog, before);
        assert!(
            catalog
                .replace("missing", entry("Unused"))
                .unwrap()
                .is_none()
        );
        assert!(catalog.replace_at(usize::MAX, entry("Unused")).is_err());

        let duplicate = format!(
            r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Résumé"/></w:docPartPr></w:docPart><w:docPart><w:docPartPr><w:name w:val="RÉSUMÉ"/></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        let (mut ambiguous, conformance) = read(duplicate.as_bytes()).unwrap();
        assert!(ambiguous.get("résumé").is_err());
        assert!(
            ambiguous
                .rename_at(1, Name::new("Repaired").unwrap())
                .unwrap()
        );
        assert_eq!(
            ambiguous.get("résumé").unwrap().unwrap().name(),
            Some("Résumé")
        );
        assert_eq!(
            ambiguous.get("repaired").unwrap().unwrap().name(),
            Some("Repaired")
        );
        let plan = plan_write(&ambiguous, conformance).unwrap();
        let index = conformance.index();
        let mut shell = XmlSize::default();
        write_catalog_open(&mut shell, conformance).unwrap();
        shell.push_str("<w:docParts></w:docParts>").unwrap();
        shell.push_str("</w:glossaryDocument>").unwrap();
        assert_eq!(
            plan.bytes,
            shell.bytes
                + ambiguous.state.background_bytes[index]
                + ambiguous.state.entry_bytes[index]
        );
    }

    #[test]
    fn empty_producer_names_and_categories_survive_unrelated_crud() {
        let xml = format!(
            r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val=""/><w:category><w:name w:val=""/><w:gallery w:val="autoTxt"/></w:category></w:docPartPr></w:docPart><w:docPart><w:docPartPr><w:name w:val="  "/><w:category><w:name w:val="  "/><w:gallery w:val="autoTxt"/></w:category></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        let (mut catalog, conformance) = read(xml.as_bytes()).unwrap();
        assert_eq!(catalog.at(0).unwrap().name(), Some(""));
        assert_eq!(catalog.at(1).unwrap().name(), Some("  "));
        catalog.add(entry("Authored")).unwrap();
        let output = String::from_utf8(write(&catalog, conformance).unwrap()).unwrap();
        assert!(output.contains(r#"<w:name w:val=""/>"#), "{output}");
        assert!(output.contains(r#"<w:name w:val="  "/>"#), "{output}");
        assert!(Name::new("").is_err());
        assert!(Category::new(" ", Gallery::new("autoTxt").unwrap()).is_err());
    }

    #[test]
    fn carriage_returns_remain_distinct_from_line_feeds() {
        let body = format!(
            r#"<w:docPartBody xmlns:w="{W}"><w:p><w:r><w:t>A&#xD;B
C</w:t></w:r></w:p></w:docPartBody>"#
        );
        let mut catalog = Catalog::new();
        catalog
            .add(Entry::new("Line endings", body.into_bytes()).unwrap())
            .unwrap();
        let output =
            String::from_utf8(write(&catalog, Conformance::Transitional).unwrap()).unwrap();
        assert!(output.contains("A&#xD;B\nC"), "{output}");
        let (round_trip, _) = read(output.as_bytes()).unwrap();
        let body = std::str::from_utf8(round_trip.at(0).unwrap().body().unwrap()).unwrap();
        assert!(body.contains("A&#xD;B\nC"), "{body}");

        let forbidden = format!(
            r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="bad&#x0;name"/></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        assert!(read(forbidden.as_bytes()).is_err());
    }

    #[test]
    fn opaque_namespace_mapping_does_not_rewrite_literal_values() {
        let literal = W;
        let body = format!(
            r#"<w:docPartBody xmlns:w="{W}" xmlns:u="urn:test"><w:p u:value="prefix-{literal}-suffix"><w:r><w:t>{literal}</w:t></w:r></w:p></w:docPartBody>"#
        );
        let catalog = Catalog {
            background: None,
            background_refs: Arc::from([]),
            background_lineage: None,
            entries: vec![Entry::new("Literal", body.into_bytes()).unwrap()],
            binding: None,
            state: CatalogState::default(),
        };
        let strict = String::from_utf8(write(&catalog, Conformance::Strict).unwrap()).unwrap();
        assert!(strict.contains(&format!(r#"u:value="prefix-{W}-suffix""#)));
        assert!(strict.contains(&format!(">{W}</w:t>")), "{strict}");
        assert!(strict.contains(WS));
        read(strict.as_bytes()).unwrap();
    }

    #[test]
    fn strict_mce_and_malformed_input_are_bounded() {
        let xml = format!(
            r#"<w:glossaryDocument xmlns:w="{WS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:u" mc:Ignorable="u"><w:docParts><mc:AlternateContent><mc:Choice Requires="u"><u:x/></mc:Choice><mc:Fallback><w:docPart><w:docPartPr><w:name w:val="MCE"/><w:behaviors><w:behavior w:val="p"/></w:behaviors></w:docPartPr><w:docPartBody><w:p/></w:docPartBody></w:docPart></mc:Fallback></mc:AlternateContent></w:docParts></w:glossaryDocument>"#
        );
        assert_eq!(read(xml.as_bytes()).unwrap().0.len(), 1);
        for bad in [
            format!(
                r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:behaviors><w:behavior w:val="run"/></w:behaviors></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
            ),
            format!(r#"<!DOCTYPE x><w:glossaryDocument xmlns:w="{W}"/>"#),
        ] {
            assert!(read(bad.as_bytes()).is_err(), "accepted {bad}");
        }
    }

    #[test]
    fn producer_optional_metadata_is_readable_and_retained() {
        let xml = format!(
            r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:style w:val="ProducerStyle"/><w:guid/></w:docPartPr><w:docPartBody><w:p/></w:docPartBody></w:docPart><w:docPart><w:docPartBody><w:p/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        let (mut catalog, conformance) = read(xml.as_bytes()).unwrap();
        assert_eq!(catalog.len(), 2);
        assert!(catalog.at(0).unwrap().name().is_none());
        assert!(catalog.at(0).unwrap().props().unwrap().id.is_none());
        assert!(catalog.at(1).unwrap().props().is_none());

        catalog.add(entry("Authored")).unwrap();
        let rewritten = String::from_utf8(write(&catalog, conformance).unwrap()).unwrap();
        assert!(rewritten.contains("<w:guid/>"), "{rewritten}");
        assert_eq!(read(rewritten.as_bytes()).unwrap().0.len(), 3);
    }

    #[test]
    fn unrelated_crud_retains_ignorable_producer_content() {
        let xml = format!(
            r#"<w:glossaryDocument xmlns:w="{W}" xmlns:mc="{MC}" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w:docParts><w:docPart mc:Ignorable="w15"><w:docPartPr><w:name w:val="Keep"/></w:docPartPr><w:docPartBody><w:p><w15:producerExtension w15:value="preserve-me"/></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        let (mut catalog, conformance) = read(xml.as_bytes()).unwrap();
        assert!(
            !std::str::from_utf8(catalog.at(0).unwrap().body().unwrap())
                .unwrap()
                .contains("producerExtension")
        );
        catalog.add(entry("Added")).unwrap();

        let rewritten = String::from_utf8(write(&catalog, conformance).unwrap()).unwrap();
        assert!(rewritten.contains("producerExtension"), "{rewritten}");
        assert!(rewritten.contains("mc:Ignorable=\"w15\""), "{rewritten}");
        assert_eq!(read(rewritten.as_bytes()).unwrap().0.len(), 2);

        let keep = catalog.remove("Keep").unwrap().unwrap();
        let mut props = keep.props().unwrap().clone();
        props.description = Some("changed".to_owned());
        catalog.add(keep.with_props(props).unwrap()).unwrap();
        let targeted = String::from_utf8(write(&catalog, conformance).unwrap()).unwrap();
        assert!(!targeted.contains("producerExtension"), "{targeted}");
        assert!(targeted.contains(r#"w:description w:val="changed""#));
    }

    #[test]
    fn strict_rejects_transitional_relationships_lexicals_and_vml() {
        for xml in [
            format!(
                r#"<w:glossaryDocument xmlns:w="{WS}" xmlns:r="{R}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Mixed"/></w:docPartPr><w:docPartBody><w:p r:embed="rId1"/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#
            ),
            format!(
                r#"<w:glossaryDocument xmlns:w="{WS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Lexical" w:decorated="on"/></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
            ),
            format!(
                r#"<w:glossaryDocument xmlns:w="{WS}" xmlns:v="{VML}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="VML"/></w:docPartPr><w:docPartBody><v:shape/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#
            ),
        ] {
            assert!(read(xml.as_bytes()).is_err(), "accepted {xml}");
        }

        let transitional_vml = format!(
            r#"<w:glossaryDocument xmlns:w="{W}" xmlns:v="{VML}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="VML"/></w:docPartPr><w:docPartBody><v:shape/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        let (catalog, conformance) = read(transitional_vml.as_bytes()).unwrap();
        assert_eq!(conformance, Conformance::Transitional);
        assert!(write(&catalog, Conformance::Transitional).is_ok());
        assert!(write(&catalog, Conformance::Strict).is_err());
    }

    #[test]
    fn inherited_namespace_scope_is_emitted_once_per_opaque_root() {
        let declarations = (0..32)
            .map(|index| format!(r#" xmlns:u{index}="urn:{index}:{}""#, "x".repeat(256)))
            .collect::<String>();
        let children = "<w:r/>".repeat(1_000);
        let xml = format!(
            r#"<w:glossaryDocument xmlns:w="{W}"{declarations}><w:docParts><w:docPart><w:docPartPr><w:name w:val="Bounded"/></w:docPartPr><w:docPartBody><w:p>{children}</w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        let (catalog, _) = read(xml.as_bytes()).unwrap();
        let body = catalog.at(0).unwrap().body().unwrap();
        assert!(
            body.len() < 32 * 1024,
            "opaque body grew to {} bytes",
            body.len()
        );
    }

    #[test]
    fn dense_common_prefix_mce_scope_merges_linearly() {
        let common = "a".repeat(256);
        let attribute = |index: usize, value: &str| {
            let local = format!("{common}{index:04}");
            Attr {
                q: format!("mc:{local}"),
                ns: Arc::from(MC),
                l: local,
                v: value.to_owned(),
            }
        };
        let mut output = (0..2_048)
            .map(|index| attribute(index, "old"))
            .collect::<Vec<_>>();
        let node = Node {
            q: "w:docParts".to_owned(),
            ns: Arc::from(W),
            l: "docParts".to_owned(),
            attrs: (0..2_048)
                .rev()
                .map(|index| attribute(index, "new"))
                .collect(),
            bindings: Arc::new(NamespaceFrame::default()),
            content: Vec::new(),
        };

        merge_mce_attributes(&mut output, &node).unwrap();
        assert_eq!(output.len(), 2_048);
        assert!(output.iter().all(|attribute| attribute.v == "new"));
    }

    #[test]
    fn aggregate_projection_and_dense_content_are_bounded() {
        let declarations = (0..512)
            .map(|index| format!(r#" xmlns:u{index}="urn:{index}:{}""#, "x".repeat(128)))
            .collect::<String>();
        let entries = (0..600)
            .map(|index| {
                format!(
                    r#"<w:docPart><w:docPartPr><w:name w:val="Entry {index}"/></w:docPartPr><w:docPartBody><w:p/></w:docPartBody></w:docPart>"#
                )
            })
            .collect::<String>();
        let amplified = format!(
            r#"<w:glossaryDocument xmlns:w="{W}"{declarations}><w:docParts>{entries}</w:docParts></w:glossaryDocument>"#
        );
        assert!(amplified.len() < MAX);
        let root = parse_dom(amplified.as_bytes()).unwrap();
        let error = project(&root).unwrap_err().to_string();
        assert!(error.contains("semantic XML exceeds"), "{error}");

        let dense = "x<!--c-->".repeat((MAX_DOM_CONTENT / 2) + 1);
        let dense = format!(
            r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Dense"/></w:docPartPr><w:docPartBody><w:p>{dense}</w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        assert!(dense.len() < MAX);
        assert!(read(dense.as_bytes()).is_err());
    }

    #[test]
    fn exact_noop_preserves_custom_root_bytes_and_signature() {
        let mut package = package();
        let mut catalog = Catalog::new();
        catalog.add(entry("Keep")).unwrap();
        let source = Arc::new(write(&catalog, Conformance::Transitional).unwrap());
        let custom = PackURI::new("/word/glossary/custom.xml").unwrap();
        package
            .try_add_part(Box::new(BlobPart::new_shared(
                custom.clone(),
                CT.to_owned(),
                Arc::clone(&source),
            )))
            .unwrap();
        let main = package.main_document_part().unwrap().partname().clone();
        package
            .get_part_mut(&main)
            .unwrap()
            .rels_mut()
            .get_or_add(REL, "glossary/custom.xml");
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                Vec::new(),
            )))
            .unwrap();
        package.rels_mut().add_relationship(
            rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rSignature".to_owned(),
            false,
        );
        assert!(package.is_signed());
        let (loaded, conformance) = load(&package).unwrap().unwrap();
        assert!(!put(&mut package, loaded, conformance).unwrap());
        assert_eq!(package.get_part(&custom).unwrap().blob(), source.as_slice());
        assert!(package.is_signed());
        assert_eq!(
            raw::load(&package).unwrap().unwrap().root_name(),
            custom.as_str()
        );
    }

    #[test]
    fn semantic_create_allocates_around_unrelated_canonical_parts() {
        let mut package = package();
        let occupied = [
            "/word/glossary/document.xml",
            "/word/glossary/styles.xml",
            "/word/glossary/settings.xml",
            "/word/glossary/fontTable.xml",
            "/word/glossary/webSettings.xml",
        ];
        for (index, name) in occupied.iter().enumerate() {
            package
                .try_add_part(Box::new(BlobPart::new(
                    PackURI::new(*name).unwrap(),
                    "application/vnd.example.unrelated+xml".to_owned(),
                    format!("<unrelated id=\"{index}\"/>").into_bytes(),
                )))
                .unwrap();
        }

        let mut catalog = Catalog::new();
        catalog.add(entry("Allocated")).unwrap();
        assert!(put(&mut package, catalog, Conformance::Transitional).unwrap());

        let graph = raw::load(&package).unwrap().unwrap();
        assert_eq!(graph.root_name(), "/word/glossary/document1.xml");
        let names = graph
            .parts
            .iter()
            .map(|part| part.name.as_str())
            .collect::<HashSet<_>>();
        assert_eq!(
            names,
            HashSet::from([
                "/word/glossary/styles1.xml",
                "/word/glossary/settings1.xml",
                "/word/glossary/fontTable1.xml",
                "/word/glossary/webSettings1.xml",
            ])
        );

        raw::remove(&mut package).unwrap().unwrap();
        for (index, name) in occupied.iter().enumerate() {
            assert_eq!(
                package
                    .get_part(&PackURI::new(*name).unwrap())
                    .unwrap()
                    .blob(),
                format!("<unrelated id=\"{index}\"/>").as_bytes()
            );
        }
    }

    #[test]
    fn repeated_changed_raw_graph_is_a_signature_preserving_noop() {
        let mut package = package();
        let mut initial = Catalog::new();
        initial.add(entry("Initial")).unwrap();
        raw::put(
            &mut package,
            &raw::Graph::new(initial, Conformance::Transitional),
        )
        .unwrap();

        let mut changed = raw::load(&package).unwrap().unwrap();
        changed.catalog.add(entry("Changed")).unwrap();
        assert!(raw::put(&mut package, &changed).unwrap());
        mark_signed(&mut package);
        assert!(package.is_signed());

        assert!(!raw::put(&mut package, &changed).unwrap());
        assert!(package.is_signed());
        assert!(
            raw::load(&package)
                .unwrap()
                .unwrap()
                .catalog
                .get("changed")
                .unwrap()
                .is_some()
        );
    }

    #[test]
    fn signature_namespace_auxiliary_is_rejected_atomically() {
        let mut package = package();
        let mut existing = Catalog::new();
        existing.add(entry("Keep")).unwrap();
        raw::put(
            &mut package,
            &raw::Graph::new(existing, Conformance::Transitional),
        )
        .unwrap();
        mark_signed(&mut package);

        let mut invalid = raw::load(&package).unwrap().unwrap();
        invalid.rels.push(raw::Rel {
            id: "rIdSignatureImage".to_owned(),
            kind: rt::IMAGE.to_owned(),
            target: "/_XMLSIGNATURES/glossary.png".to_owned(),
            external: false,
        });
        assert!(
            raw::Part::new(
                "/_XMLSIGNATURES/glossary.png",
                "image/png",
                vec![0x89, b'P', b'N', b'G'],
            )
            .is_err()
        );
        invalid.parts.push(raw::Part::from_shared(
            "/_XMLSIGNATURES/glossary.png".to_owned(),
            "image/png".to_owned(),
            Arc::new(vec![0x89, b'P', b'N', b'G']),
            Vec::new(),
        ));

        assert!(raw::put(&mut package, &invalid).is_err());
        assert!(package.is_signed());
        assert!(
            raw::load(&package)
                .unwrap()
                .unwrap()
                .catalog
                .get("keep")
                .unwrap()
                .is_some()
        );
    }

    #[test]
    fn reserved_parts_and_invalid_raw_metadata_are_rejected_atomically() {
        assert!(raw::Part::new("/[Content_Types].xml", "application/xml", Vec::new()).is_err());
        assert!(
            raw::Part::new(
                "/word/glossary/media/image.png",
                "image/png;broken",
                Vec::new()
            )
            .is_err()
        );
        for (index, content_type) in [
            ct::OPC_DIGITAL_SIGNATURE_ORIGIN,
            ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE,
            ct::OPC_DIGITAL_SIGNATURE_CERTIFICATE,
            ct::OPC_RELATIONSHIPS,
        ]
        .into_iter()
        .enumerate()
        {
            assert!(
                raw::Part::new(
                    format!("/word/glossary/reserved-{index}.bin"),
                    content_type,
                    Vec::new(),
                )
                .is_err()
            );

            let mut candidate = package();
            mark_signed(&mut candidate);
            let mut graph = raw::Graph::new(Catalog::new(), Conformance::Transitional);
            graph.rels.push(raw::Rel {
                id: "rIdReservedType".to_owned(),
                kind: rt::IMAGE.to_owned(),
                target: "reserved.bin".to_owned(),
                external: false,
            });
            graph.parts.push(raw::Part::from_shared(
                "/word/glossary/reserved.bin".to_owned(),
                content_type.to_owned(),
                Arc::new(Vec::new()),
                Vec::new(),
            ));
            assert!(raw::put(&mut candidate, &graph).is_err());
            assert!(candidate.is_signed());
            assert!(raw::load(&candidate).unwrap().is_none());
        }

        let mut package = package();
        mark_signed(&mut package);
        let mut reserved = raw::Graph::new(Catalog::new(), Conformance::Transitional);
        reserved.rels.push(raw::Rel {
            id: "rIdImage".to_owned(),
            kind: rt::IMAGE.to_owned(),
            target: "/[Content_Types].xml".to_owned(),
            external: false,
        });
        reserved.parts.push(raw::Part::from_shared(
            "/[Content_Types].xml".to_owned(),
            "image/png".to_owned(),
            Arc::new(vec![0x89, b'P', b'N', b'G']),
            Vec::new(),
        ));
        assert!(raw::put(&mut package, &reserved).is_err());
        assert!(package.is_signed());
        assert!(raw::load(&package).unwrap().is_none());

        let mut invalid_id = raw::Graph::new(Catalog::new(), Conformance::Transitional);
        invalid_id.rels.push(raw::Rel {
            id: "bad id".to_owned(),
            kind: rt::HYPERLINK.to_owned(),
            target: "https://example.invalid".to_owned(),
            external: true,
        });
        assert!(raw::put(&mut package, &invalid_id).is_err());
        assert!(package.is_signed());
    }

    #[test]
    fn graph_wide_relationship_limit_is_failure_atomic() {
        let mut graph = raw::Graph::new(Catalog::new(), Conformance::Transitional);
        for part_index in 0..2 {
            let name = format!("numbering{part_index}.xml");
            graph.rels.push(raw::Rel {
                id: format!("rIdNumbering{part_index}"),
                kind: rt::NUMBERING.to_owned(),
                target: name.clone(),
                external: false,
            });
            let mut part = raw::Part::new(
                format!("/word/glossary/{name}"),
                ct::WML_NUMBERING,
                format!(r#"<w:numbering xmlns:w="{W}"/>"#).into_bytes(),
            )
            .unwrap();
            for relationship_index in 0..(MAX_VALUES / 2) {
                part.rels.push(raw::Rel {
                    id: format!("rIdLink{relationship_index}"),
                    kind: rt::HYPERLINK.to_owned(),
                    target: format!("https://example.invalid/{part_index}/{relationship_index}"),
                    external: true,
                });
            }
            graph.parts.push(part);
        }

        let mut package = package();
        mark_signed(&mut package);
        let error = raw::put(&mut package, &graph).unwrap_err().to_string();
        assert!(error.contains("graph-wide relationship limit"), "{error}");
        assert!(package.is_signed());
        assert!(raw::load(&package).unwrap().is_none());
    }

    #[test]
    fn shared_inbound_owned_parts_block_changed_store_and_remove() {
        let mut package = package();
        let mut catalog = Catalog::new();
        catalog.add(entry("Keep")).unwrap();
        put(&mut package, catalog.clone(), Conformance::Transitional).unwrap();
        let root = PackURI::new("/word/glossary/document.xml").unwrap();
        let mut referrer = BlobPart::new(
            PackURI::new("/word/header1.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml".into(),
            format!(r#"<w:hdr xmlns:w="{W}"/>"#).into_bytes(),
        );
        referrer
            .rels_mut()
            .get_or_add(rt::IMAGE, "glossary/document.xml");
        package.try_add_part(Box::new(referrer)).unwrap();

        let mut changed = catalog;
        changed.add(entry("Changed")).unwrap();
        assert!(put(&mut package, changed, Conformance::Transitional).is_err());
        assert!(remove(&mut package).is_err());
        assert!(package.get_part(&root).is_ok());
    }

    #[test]
    fn auxiliary_relationships_must_match_the_package_conformance_family() {
        for (conformance, wrong_kind) in [
            (Conformance::Transitional, rt::STRICT_IMAGE),
            (Conformance::Strict, rt::IMAGE),
        ] {
            let mut package = package_for(conformance);
            let mut graph = raw::Graph::new(Catalog::new(), conformance);
            graph.rels.push(raw::Rel {
                id: "rIdImage".to_owned(),
                kind: wrong_kind.to_owned(),
                target: "media/image.bin".to_owned(),
                external: false,
            });

            let error = raw::put(&mut package, &graph).unwrap_err().to_string();
            assert!(error.contains("conformance"), "{error}");
            assert!(raw::load(&package).unwrap().is_none());
        }
    }

    #[test]
    fn relationship_targets_use_exact_content_type_profiles() {
        let mut invalid_package = package();
        let mut invalid = raw::Graph::new(Catalog::new(), Conformance::Transitional);
        invalid.rels.push(raw::Rel {
            id: "rIdStyles".to_owned(),
            kind: rt::STYLES.to_owned(),
            target: "styles.png".to_owned(),
            external: false,
        });
        invalid.parts.push(
            raw::Part::new(
                "/word/glossary/styles.png",
                "image/png",
                vec![0x89, b'P', b'N', b'G'],
            )
            .unwrap(),
        );
        assert!(raw::put(&mut invalid_package, &invalid).is_err());
        assert!(raw::load(&invalid_package).unwrap().is_none());

        let mut valid_package = package();
        let mut valid = raw::Graph::new(Catalog::new(), Conformance::Transitional);
        valid.rels.push(raw::Rel {
            id: "rIdCustomizations".to_owned(),
            kind: CUSTOMIZATIONS_REL.to_owned(),
            target: "customizations.xml".to_owned(),
            external: false,
        });
        valid.parts.push(
            raw::Part::new(
                "/word/glossary/customizations.xml",
                CUSTOMIZATIONS_CT,
                Vec::new(),
            )
            .unwrap(),
        );
        assert!(raw::put(&mut valid_package, &valid).unwrap());
        assert_eq!(raw::load(&valid_package).unwrap().unwrap().parts.len(), 1);
    }

    #[test]
    fn standards_relationship_matrix_round_trips_owned_parts_and_keeps_references() {
        fn relationship(id: &str, suffix: &str, target: &str, external: bool) -> raw::Rel {
            raw::Rel {
                id: id.to_owned(),
                kind: format!("{R}/{suffix}"),
                target: target.to_owned(),
                external,
            }
        }

        let body = format!(
            r#"<w:docPartBody xmlns:w="{W}" xmlns:r="{R}"><w:p r:dm="rIdDiagram"/></w:docPartBody>"#
        );
        let mut catalog = Catalog::new();
        catalog
            .add(Entry::new("Matrix", body.into_bytes()).unwrap())
            .unwrap();
        let mut graph = raw::Graph::new(catalog, Conformance::Transitional);
        graph.rels = vec![
            relationship("rIdComments", "comments", "comments.xml", false),
            relationship("rIdSettings", "settings", "settings.xml", false),
            relationship("rIdFonts", "fontTable", "fontTable.xml", false),
            relationship("rIdWeb", "webSettings", "webSettings.xml", false),
            relationship("rIdDiagram", "diagramData", "diagram/data.xml", false),
            relationship("rIdCustomXml", "customXml", "customXml/item1.xml", false),
            relationship("rIdControl", "control", "activeX/activeX1.xml", false),
            relationship("rIdGenericControl", "control", "controls/vendor.bin", false),
            relationship("rIdChunk", "aFChunk", "chunks/chunk.html", false),
            relationship(
                "rIdReference",
                "hyperlink",
                "../../shared/reference.xml",
                false,
            ),
            relationship(
                "rIdExternalImage",
                "image",
                "https://example.invalid/image.png",
                true,
            ),
            raw::Rel {
                id: "rIdCustomizations".to_owned(),
                kind: CUSTOMIZATIONS_REL.to_owned(),
                target: "customizations.xml".to_owned(),
                external: false,
            },
        ];

        let mut comments = raw::Part::new(
            "/word/glossary/comments.xml",
            ct::WML_COMMENTS,
            format!(r#"<w:comments xmlns:w="{W}"/>"#).into_bytes(),
        )
        .unwrap();
        comments.rels = vec![
            relationship("rIdChart", "chart", "charts/chart1.xml", false),
            relationship(
                "rIdVideo",
                "video",
                "https://example.invalid/video.mp4",
                true,
            ),
        ];
        let mut chart = raw::Part::new(
            "/word/glossary/charts/chart1.xml",
            ct::DML_CHART,
            b"<c:chartSpace xmlns:c=\"urn:chart\"/>".to_vec(),
        )
        .unwrap();
        chart.rels.push(relationship(
            "rIdPackage",
            "package",
            "https://example.invalid/data.xlsx",
            true,
        ));
        chart.rels.extend([
            raw::Rel {
                id: "rIdChartStyle".to_owned(),
                kind: CHART_STYLE_REL.to_owned(),
                target: "style1.xml".to_owned(),
                external: false,
            },
            raw::Rel {
                id: "rIdChartColors".to_owned(),
                kind: CHART_COLOR_STYLE_REL.to_owned(),
                target: "colors1.xml".to_owned(),
                external: false,
            },
            relationship(
                "rIdChartShapes",
                "chartUserShapes",
                "userShapes1.xml",
                false,
            ),
            raw::Rel {
                id: "rIdChartStyle2012".to_owned(),
                kind: CHART_STYLE_REL_2012.to_owned(),
                target: "style1.xml".to_owned(),
                external: false,
            },
            raw::Rel {
                id: "rIdChartColors2012".to_owned(),
                kind: CHART_COLOR_STYLE_REL_2012.to_owned(),
                target: "colors1.xml".to_owned(),
                external: false,
            },
            relationship(
                "rIdThemeOverride",
                "themeOverride",
                "themeOverride1.xml",
                false,
            ),
        ]);
        let mut chart_shapes = raw::Part::new(
            "/word/glossary/charts/userShapes1.xml",
            ct::DML_CHARTSHAPES,
            Vec::new(),
        )
        .unwrap();
        chart_shapes.rels.push(relationship(
            "rIdLinkedImage",
            "image",
            "https://example.invalid/chart-image.png",
            true,
        ));
        chart_shapes.rels.extend([
            relationship("rIdChartBack", "chart", "chart1.xml", false),
            relationship(
                "rIdInkCustomXml",
                "customXml",
                "../customXml/item1.xml",
                false,
            ),
        ]);
        let mut custom_xml = raw::Part::new(
            "/word/glossary/customXml/item1.xml",
            "application/xml",
            b"<data/>".to_vec(),
        )
        .unwrap();
        custom_xml.rels.push(relationship(
            "rIdCustomXmlProps",
            "customXmlProps",
            "itemProps1.xml",
            false,
        ));
        let mut control = raw::Part::new(
            "/word/glossary/activeX/activeX1.xml",
            ACTIVE_X_DESCRIPTOR_CT,
            b"<ax:ocx xmlns:ax=\"urn:active-x\"/>".to_vec(),
        )
        .unwrap();
        control.rels.push(raw::Rel {
            id: "rIdControlBinary".to_owned(),
            kind: ACTIVE_X_BINARY_REL.to_owned(),
            target: "activeX1.bin".to_owned(),
            external: false,
        });
        let mut settings = raw::Part::new(
            "/word/glossary/settings.xml",
            ct::WML_SETTINGS,
            format!(r#"<w:settings xmlns:w="{W}"/>"#).into_bytes(),
        )
        .unwrap();
        settings.rels.push(relationship(
            "rIdTemplate",
            "attachedTemplate",
            "https://example.invalid/template.dotx",
            true,
        ));
        settings.rels.push(relationship(
            "rIdRecipientData",
            "recipientData",
            "recipientData.xml",
            false,
        ));
        let mut font_table = raw::Part::new(
            "/word/glossary/fontTable.xml",
            ct::WML_FONT_TABLE,
            format!(r#"<w:fonts xmlns:w="{W}"/>"#).into_bytes(),
        )
        .unwrap();
        font_table
            .rels
            .push(relationship("rIdFont", "font", "fonts/font.ttf", false));
        let mut web_settings = raw::Part::new(
            "/word/glossary/webSettings.xml",
            ct::WML_WEB_SETTINGS,
            format!(r#"<w:webSettings xmlns:w="{W}"/>"#).into_bytes(),
        )
        .unwrap();
        web_settings.rels.push(relationship(
            "rIdFrame",
            "frame",
            "https://example.invalid/frame.html",
            true,
        ));
        let mut diagram_data = raw::Part::new(
            "/word/glossary/diagram/data.xml",
            ct::DML_DIAGRAM_DATA,
            b"<d:dataModel xmlns:d=\"urn:diagram\"/>".to_vec(),
        )
        .unwrap();
        diagram_data.rels.push(raw::Rel {
            id: "rIdDrawing".to_owned(),
            kind: DIAGRAM_DRAWING_REL.to_owned(),
            target: "drawing.xml".to_owned(),
            external: false,
        });
        diagram_data.rels.push(relationship(
            "rIdDiagramLink",
            "hyperlink",
            "https://example.invalid/diagram",
            true,
        ));
        let mut customizations = raw::Part::new(
            "/word/glossary/customizations.xml",
            CUSTOMIZATIONS_CT,
            Vec::new(),
        )
        .unwrap();
        customizations.rels.push(raw::Rel {
            id: "rIdToolbars".to_owned(),
            kind: ATTACHED_TOOLBARS_REL.to_owned(),
            target: "attachedToolbars.bin".to_owned(),
            external: false,
        });
        graph.parts = vec![
            comments,
            chart,
            raw::Part::new(
                "/word/glossary/charts/style1.xml",
                CHART_STYLE_CT,
                Vec::new(),
            )
            .unwrap(),
            raw::Part::new(
                "/word/glossary/charts/colors1.xml",
                CHART_COLOR_STYLE_CT,
                Vec::new(),
            )
            .unwrap(),
            raw::Part::new(
                "/word/glossary/charts/themeOverride1.xml",
                ct::OFC_THEME_OVERRIDE,
                Vec::new(),
            )
            .unwrap(),
            chart_shapes,
            custom_xml,
            raw::Part::new(
                "/word/glossary/customXml/itemProps1.xml",
                litchi_ooxml_common::custom_xml::PROPS_CONTENT_TYPE,
                b"<ds:datastoreItem xmlns:ds=\"urn:custom-xml-props\"/>".to_vec(),
            )
            .unwrap(),
            control,
            raw::Part::new(
                "/word/glossary/activeX/activeX1.bin",
                ACTIVE_X_BINARY_CT,
                vec![0, 1, 2],
            )
            .unwrap(),
            raw::Part::new(
                "/word/glossary/controls/vendor.bin",
                "application/vnd.example.control",
                vec![9, 8, 7],
            )
            .unwrap(),
            settings,
            raw::Part::new(
                "/word/glossary/recipientData.xml",
                RECIPIENT_DATA_CT,
                Vec::new(),
            )
            .unwrap(),
            font_table,
            web_settings,
            diagram_data,
            raw::Part::new(
                "/word/glossary/diagram/drawing.xml",
                ct::DML_DIAGRAM_DRAWING,
                Vec::new(),
            )
            .unwrap(),
            raw::Part::new(
                "/word/glossary/chunks/chunk.html",
                "text/html",
                b"<p>chunk</p>".to_vec(),
            )
            .unwrap(),
            raw::Part::new("/word/glossary/fonts/font.ttf", FONT_TTF_CT, vec![0, 1, 2]).unwrap(),
            customizations,
            raw::Part::new(
                "/word/glossary/attachedToolbars.bin",
                ATTACHED_TOOLBARS_CT,
                Vec::new(),
            )
            .unwrap(),
        ];

        let mut package = package();
        let reference = PackURI::new("/shared/reference.xml").unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                reference.clone(),
                "application/xml".to_owned(),
                b"<shared/>".to_vec(),
            )))
            .unwrap();
        assert!(raw::put(&mut package, &graph).unwrap());
        let loaded = raw::load(&package).unwrap().unwrap();
        assert!(
            !loaded
                .parts
                .iter()
                .any(|part| part.name == reference.as_str())
        );
        let removed = raw::remove(&mut package).unwrap().unwrap();
        assert_eq!(removed.parts.len(), graph.parts.len());
        assert_eq!(package.get_part(&reference).unwrap().blob(), b"<shared/>");
    }

    #[test]
    fn relationship_roles_reject_invalid_modes_names_and_content_type_spoofing() {
        for (kind, target, external) in [
            (format!("{R}/theme"), "theme.xml", false),
            (format!("{R}/afChunk"), "chunk.html", false),
            (
                format!("{R}/styles"),
                "https://example.invalid/styles",
                true,
            ),
        ] {
            let mut graph = raw::Graph::new(Catalog::new(), Conformance::Transitional);
            graph.rels.push(raw::Rel {
                id: "rIdInvalid".to_owned(),
                kind,
                target: target.to_owned(),
                external,
            });
            if !external {
                graph.parts.push(
                    raw::Part::new(
                        format!("/word/glossary/{target}"),
                        if target == "theme.xml" {
                            ct::OFC_THEME
                        } else {
                            "text/html"
                        },
                        Vec::new(),
                    )
                    .unwrap(),
                );
            }
            assert!(raw::put(&mut package(), &graph).is_err());
        }

        let mut spoofed = raw::Graph::new(Catalog::new(), Conformance::Transitional);
        spoofed.rels.push(raw::Rel {
            id: "rIdPackage".to_owned(),
            kind: format!("{R}/package"),
            target: "embedded.bin".to_owned(),
            external: false,
        });
        let mut embedded =
            raw::Part::new("/word/glossary/embedded.bin", ct::DML_CHART, Vec::new()).unwrap();
        embedded.rels.push(raw::Rel {
            id: "rIdShapes".to_owned(),
            kind: format!("{R}/chartUserShapes"),
            target: "shapes.xml".to_owned(),
            external: false,
        });
        spoofed.parts = vec![
            embedded,
            raw::Part::new("/word/glossary/shapes.xml", ct::DML_CHARTSHAPES, Vec::new()).unwrap(),
        ];
        assert!(raw::put(&mut package(), &spoofed).is_err());

        fn child_graph(
            root_kind: &str,
            parent_content_type: &str,
            child_kind: &str,
            child_content_type: Option<&str>,
            external: bool,
        ) -> raw::Graph {
            let mut graph = raw::Graph::new(Catalog::new(), Conformance::Transitional);
            graph.rels.push(raw::Rel {
                id: "rIdParent".to_owned(),
                kind: root_kind.to_owned(),
                target: "parent.xml".to_owned(),
                external: false,
            });
            let mut parent =
                raw::Part::new("/word/glossary/parent.xml", parent_content_type, Vec::new())
                    .unwrap();
            parent.rels.push(raw::Rel {
                id: "rIdChild".to_owned(),
                kind: child_kind.to_owned(),
                target: if external {
                    "https://example.invalid/child".to_owned()
                } else {
                    "child.bin".to_owned()
                },
                external,
            });
            graph.parts.push(parent);
            if let Some(content_type) = child_content_type {
                graph.parts.push(
                    raw::Part::new("/word/glossary/child.bin", content_type, Vec::new()).unwrap(),
                );
            }
            graph
        }

        for invalid in [
            child_graph(
                &format!("{R}/customXml"),
                "application/xml",
                &format!("{R}/customXmlProps"),
                Some("application/xml"),
                false,
            ),
            child_graph(
                &format!("{R}/chart"),
                ct::DML_CHART,
                CHART_STYLE_REL,
                Some("application/xml"),
                false,
            ),
            child_graph(
                &format!("{R}/control"),
                ACTIVE_X_DESCRIPTOR_CT,
                ACTIVE_X_BINARY_REL,
                Some(ACTIVE_X_BINARY_CT),
                true,
            ),
            child_graph(
                &format!("{R}/control"),
                "application/vnd.example.control",
                ACTIVE_X_BINARY_REL,
                Some(ACTIVE_X_BINARY_CT),
                false,
            ),
            child_graph(
                &format!("{R}/customXml"),
                "application/xml",
                &format!("{R}/hyperlink"),
                None,
                true,
            ),
        ] {
            assert!(raw::put(&mut package(), &invalid).is_err());
        }
    }

    #[test]
    fn semantic_catalogs_cannot_rebind_colliding_relationship_ids() {
        fn picture_graph(target: &str, payload: &[u8]) -> raw::Graph {
            let body = format!(
                r#"<w:docPartBody xmlns:w="{W}" xmlns:r="{R}"><w:p><w:r><w:drawing r:embed="rIdImage"/></w:r></w:p></w:docPartBody>"#
            );
            let mut catalog = Catalog::new();
            catalog
                .add(Entry::new("Picture", body.into_bytes()).unwrap())
                .unwrap();
            let mut graph = raw::Graph::new(catalog, Conformance::Transitional);
            graph.rels.push(raw::Rel {
                id: "rIdImage".to_owned(),
                kind: rt::IMAGE.to_owned(),
                target: format!("media/{target}.png"),
                external: false,
            });
            graph.parts.push(
                raw::Part::new(
                    format!("/word/glossary/media/{target}.png"),
                    "image/png",
                    payload.to_vec(),
                )
                .unwrap(),
            );
            graph
        }

        let source_graph = picture_graph("source", b"source image");
        let destination_graph = picture_graph("destination", b"destination image");
        let mut source = package();
        let mut destination = package();
        raw::put(&mut source, &source_graph).unwrap();
        raw::put(&mut destination, &destination_graph).unwrap();

        let (mut source_catalog, _) = load(&source).unwrap().unwrap();
        let source_entry = source_catalog.remove("picture").unwrap().unwrap();
        let (mut destination_catalog, _) = load(&destination).unwrap().unwrap();
        let before = destination_catalog.clone();
        assert!(
            destination_catalog
                .replace("picture", source_entry)
                .is_err()
        );
        assert_eq!(destination_catalog, before);

        let destination_entry = destination_catalog.remove("picture").unwrap().unwrap();
        let rebound_body = format!(
            r#"<w:docPartBody xmlns:w="{W}" xmlns:r="{R}"><w:p r:embed="rIdImage"/><w:p/></w:docPartBody>"#
        );
        let destination_entry = destination_entry
            .with_body(rebound_body.into_bytes())
            .unwrap();
        assert!(destination_catalog.add(destination_entry).is_err());
        let background = format!(r#"<w:background xmlns:w="{W}" xmlns:r="{R}" r:id="rIdImage"/>"#);
        assert!(
            destination_catalog
                .set_background(background.into_bytes())
                .is_err()
        );
        assert!(destination_catalog.background().is_none());

        let (foreign, conformance) = load(&source).unwrap().unwrap();
        assert!(put(&mut destination, foreign, conformance).is_err());
        let retained = raw::load(&destination).unwrap().unwrap();
        assert_eq!(retained.parts[0].data(), b"destination image");
    }

    #[test]
    fn inactive_mce_relationship_ids_remain_bound_to_their_physical_graph() {
        let body = format!(
            r#"<w:docPartBody xmlns:w="{W}" xmlns:r="{R}" xmlns:mc="{MC}" xmlns:u="urn:unsupported" mc:Ignorable="u"><mc:AlternateContent><mc:Choice Requires="u"><w:p r:embed="rIdInactiveImage"/></mc:Choice><mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent></w:docPartBody>"#
        );
        let mut catalog = Catalog::new();
        catalog
            .add(Entry::new("Inactive reference", body.into_bytes()).unwrap())
            .unwrap();
        let mut graph = raw::Graph::new(catalog, Conformance::Transitional);
        graph.rels.push(raw::Rel {
            id: "rIdInactiveImage".to_owned(),
            kind: rt::IMAGE.to_owned(),
            target: "media/inactive.png".to_owned(),
            external: false,
        });
        graph.parts.push(
            raw::Part::new(
                "/word/glossary/media/inactive.png",
                "image/png",
                b"inactive image".to_vec(),
            )
            .unwrap(),
        );

        let mut source = package();
        raw::put(&mut source, &graph).unwrap();
        let (mut loaded, conformance) = load(&source).unwrap().unwrap();
        assert!(
            !std::str::from_utf8(loaded.at(0).unwrap().body().unwrap())
                .unwrap()
                .contains("rIdInactiveImage")
        );
        let mut destination = package();
        assert!(put(&mut destination, loaded.clone(), conformance).is_err());
        assert!(raw::load(&destination).unwrap().is_none());

        loaded.add(entry("Unrelated")).unwrap();
        assert!(put(&mut source, loaded, conformance).unwrap());
        let preserved = raw::load(&source).unwrap().unwrap();
        assert!(
            String::from_utf8(write(&preserved.catalog, conformance).unwrap())
                .unwrap()
                .contains("rIdInactiveImage")
        );
        assert!(
            preserved
                .parts
                .iter()
                .any(|part| part.data() == b"inactive image")
        );
    }

    #[test]
    fn complete_graph_follows_valid_parts_outside_the_conventional_directory() {
        let mut graph = raw::Graph::new(Catalog::new(), Conformance::Transitional);
        graph.rels.push(raw::Rel {
            id: "rIdImage".to_owned(),
            kind: rt::IMAGE.to_owned(),
            target: "../../assets/glossary.png".to_owned(),
            external: false,
        });
        graph.parts.push(
            raw::Part::new(
                "/assets/glossary.png",
                "image/png",
                vec![0x89, b'P', b'N', b'G'],
            )
            .unwrap(),
        );
        let mut package = package();
        assert!(raw::put(&mut package, &graph).unwrap());
        let removed = raw::remove(&mut package).unwrap().unwrap();
        assert_eq!(removed.parts[0].name, "/assets/glossary.png");
        assert!(
            package
                .get_part(&PackURI::new("/assets/glossary.png").unwrap())
                .is_err()
        );
    }

    #[test]
    fn raw_remove_returns_a_graph_that_raw_put_can_restore() {
        let payload = vec![0xA5; MAX + 1];
        let mut graph = raw::Graph::new(Catalog::new(), Conformance::Transitional);
        graph.rels.push(raw::Rel {
            id: "rIdImage".to_owned(),
            kind: rt::IMAGE.to_owned(),
            target: "media/large.bin".to_owned(),
            external: false,
        });
        graph
            .parts
            .push(raw::Part::new("/word/glossary/media/large.bin", "image/png", payload).unwrap());

        let mut source = package();
        assert!(raw::put(&mut source, &graph).unwrap());
        let removed = raw::remove(&mut source).unwrap().unwrap();
        let mut destination = package();
        assert!(raw::put(&mut destination, &removed).unwrap());
        assert_eq!(
            raw::load(&destination).unwrap().unwrap().parts[0]
                .data()
                .len(),
            MAX + 1
        );
    }

    #[test]
    fn raw_transfer_preserves_the_main_owner_relationship_metadata() {
        let mut source = package();
        let root = PackURI::new("/word/glossary/custom.xml").unwrap();
        source
            .try_add_part(Box::new(BlobPart::new(
                root,
                CT.to_owned(),
                write(&Catalog::new(), Conformance::Transitional).unwrap(),
            )))
            .unwrap();
        let main = source.main_document_part().unwrap().partname().clone();
        source
            .get_part_mut(&main)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                REL.to_owned(),
                "glossary/custom.xml".to_owned(),
                "rCustomGlossary".to_owned(),
                litchi_opc::TargetMode::Internal,
            )
            .unwrap();

        let graph = raw::remove(&mut source).unwrap().unwrap();
        assert_eq!(graph.owner_relationship_id(), Some("rCustomGlossary"));
        assert_eq!(graph.owner_target(), Some("glossary/custom.xml"));
        let mut destination = package();
        raw::put(&mut destination, &graph).unwrap();
        let restored = raw::load(&destination).unwrap().unwrap();
        assert_eq!(restored.owner_relationship_id(), Some("rCustomGlossary"));
        assert_eq!(restored.owner_target(), Some("glossary/custom.xml"));
    }

    #[test]
    fn raw_transfer_rebases_owner_target_and_allocates_a_free_id() {
        let mut source = package();
        let root = PackURI::new("/word/glossary/custom.xml").unwrap();
        source
            .try_add_part(Box::new(BlobPart::new(
                root,
                CT.to_owned(),
                write(&Catalog::new(), Conformance::Transitional).unwrap(),
            )))
            .unwrap();
        let source_main = source.main_document_part().unwrap().partname().clone();
        source
            .get_part_mut(&source_main)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                REL.to_owned(),
                "glossary/custom.xml".to_owned(),
                "rCustomGlossary".to_owned(),
                litchi_opc::TargetMode::Internal,
            )
            .unwrap();
        let graph = raw::remove(&mut source).unwrap().unwrap();

        let mut destination = OpcPackage::new();
        let destination_main = PackURI::new("/custom/main.xml").unwrap();
        destination
            .try_add_part(Box::new(BlobPart::new(
                destination_main.clone(),
                ct::WML_DOCUMENT_MAIN.to_owned(),
                format!(r#"<w:document xmlns:w="{W}"><w:body/></w:document>"#).into_bytes(),
            )))
            .unwrap();
        destination.relate_to("custom/main.xml", rt::OFFICE_DOCUMENT);
        destination
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/custom/unrelated.bin").unwrap(),
                "application/octet-stream".to_owned(),
                vec![1, 2, 3],
            )))
            .unwrap();
        destination
            .get_part_mut(&destination_main)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                "urn:unrelated".to_owned(),
                "unrelated.bin".to_owned(),
                "rCustomGlossary".to_owned(),
                litchi_opc::TargetMode::Internal,
            )
            .unwrap();

        assert!(raw::put(&mut destination, &graph).unwrap());
        let restored = raw::load(&destination).unwrap().unwrap();
        assert_eq!(restored.root_name(), "/word/glossary/custom.xml");
        assert_eq!(restored.owner_main_part(), Some("/custom/main.xml"));
        assert_eq!(restored.owner_target(), Some("../word/glossary/custom.xml"));
        assert_ne!(restored.owner_relationship_id(), Some("rCustomGlossary"));
        assert_eq!(
            destination
                .get_part(&destination_main)
                .unwrap()
                .rels()
                .get("rCustomGlossary")
                .unwrap()
                .reltype(),
            "urn:unrelated"
        );
        mark_signed(&mut destination);
        assert!(!raw::put(&mut destination, &graph).unwrap());
        assert!(destination.is_signed());
        let mut bytes = Vec::new();
        destination.to_stream(&mut bytes).unwrap();
        let reopened = OpcPackage::from_bytes(&bytes).unwrap();
        assert_eq!(
            raw::load(&reopened).unwrap().unwrap().owner_main_part(),
            Some("/custom/main.xml")
        );
    }

    #[test]
    fn ownership_closure_resolves_case_varied_relationship_targets() {
        let mut package = package();
        let root = add_empty_glossary(&mut package, Conformance::Transitional);
        let auxiliary = PackURI::new("/word/glossary/media/image.bin").unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                auxiliary.clone(),
                "image/png".to_owned(),
                vec![1, 2, 3],
            )))
            .unwrap();
        package
            .get_part_mut(&root)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                "/WORD/GLOSSARY/MEDIA/IMAGE.BIN".to_owned(),
                "rIdImage".to_owned(),
                litchi_opc::TargetMode::Internal,
            )
            .unwrap();

        let graph = raw::remove(&mut package).unwrap().unwrap();
        assert!(
            graph
                .parts
                .iter()
                .any(|part| part.name == auxiliary.as_str())
        );
        assert!(package.get_part(&auxiliary).is_err());
        assert!(raw::load(&package).unwrap().is_none());
    }

    #[test]
    fn unowned_parts_in_the_conventional_glossary_directory_are_preserved() {
        for with_root in [false, true] {
            let mut package = package();
            if with_root {
                add_empty_glossary(&mut package, Conformance::Transitional);
            }
            let orphan = PackURI::new("/word/glossary/media/orphan.bin").unwrap();
            package
                .try_add_part(Box::new(BlobPart::new(
                    orphan.clone(),
                    "application/octet-stream".to_owned(),
                    vec![4, 5, 6],
                )))
                .unwrap();

            assert_eq!(raw::load(&package).unwrap().is_some(), with_root);
            assert_eq!(raw::remove(&mut package).unwrap().is_some(), with_root);
            assert_eq!(package.get_part(&orphan).unwrap().blob(), [4, 5, 6]);
        }
    }

    #[test]
    fn semantic_payload_limits_fail_before_xml_parsing() {
        let oversized = vec![b'x'; MAX + 1];
        let entry_error = Entry::new("Oversized", oversized).unwrap_err().to_string();
        assert!(
            entry_error.contains("payload exceeds 32 MiB"),
            "{entry_error}"
        );

        let mut catalog = Catalog::new();
        let background_error = catalog
            .set_background(vec![b'x'; MAX + 1])
            .unwrap_err()
            .to_string();
        assert!(
            background_error.contains("payload exceeds 32 MiB"),
            "{background_error}"
        );
        assert!(catalog.background().is_none());
    }

    #[test]
    fn aggregate_catalog_budget_is_checked_atomically() {
        let mut catalog = Catalog::new();
        for index in 0..5 {
            let name = format!("Large {index}");
            let props = Props {
                description: Some("\"".repeat(MAX_STRING)),
                ..Props::new(Name::new(&name).unwrap())
            };
            catalog
                .add(entry(&name).with_props(props).unwrap())
                .unwrap();
        }
        let before = catalog.len();
        let rejected_name = "Too large";
        let rejected_props = Props {
            description: Some("\"".repeat(MAX_STRING)),
            ..Props::new(Name::new(rejected_name).unwrap())
        };
        let error = catalog
            .add(entry(rejected_name).with_props(rejected_props).unwrap())
            .unwrap_err()
            .to_string();
        assert!(error.contains("exceeds 32 MiB"), "{error}");
        assert_eq!(catalog.len(), before);
        assert!(catalog.get(rejected_name).unwrap().is_none());

        let replacement_props = Props {
            style: Some("\"".repeat(MAX_STRING)),
            description: Some("\"".repeat(MAX_STRING)),
            ..Props::new(Name::new("Large 0").unwrap())
        };
        let replacement_error = catalog
            .put(entry("Large 0").with_props(replacement_props).unwrap())
            .unwrap_err()
            .to_string();
        assert!(
            replacement_error.contains("exceeds 32 MiB"),
            "{replacement_error}"
        );
        assert!(
            catalog
                .get("Large 0")
                .unwrap()
                .unwrap()
                .props()
                .unwrap()
                .style
                .is_none()
        );

        let background = format!(
            r#"<w:background xmlns:w="{W}"><w:color>{}</w:color></w:background>"#,
            ">".repeat(600_000)
        );
        let background_error = catalog
            .set_background(background.into_bytes())
            .unwrap_err()
            .to_string();
        assert!(
            background_error.contains("exceeds 32 MiB"),
            "{background_error}"
        );
        assert!(catalog.background().is_none());
    }

    #[test]
    fn canonical_write_matches_its_checked_plan() {
        let mut catalog = Catalog::new();
        catalog.add(entry("Planned")).unwrap();
        catalog
            .set_background(
                format!(r#"<w:background xmlns:w="{W}" w:color="A&amp;B"/>"#).into_bytes(),
            )
            .unwrap();

        for conformance in [Conformance::Transitional, Conformance::Strict] {
            let plan = plan_write(&catalog, conformance).unwrap();
            let xml = write(&catalog, conformance).unwrap();
            assert_eq!(xml.len(), plan.bytes);
            assert_eq!(read(&xml).unwrap().1, conformance);
        }
    }

    #[test]
    fn aggregate_auxiliary_payload_limit_rejects_before_mutation() {
        let mut package = package();
        let payload = Arc::new(vec![0u8; MAX]);
        let mut graph = raw::Graph::new(Catalog::new(), Conformance::Transitional);
        for index in 0..9 {
            graph.parts.push(raw::Part::from_shared(
                format!("/word/glossary/media/data{index}.bin"),
                "application/octet-stream".to_owned(),
                Arc::clone(&payload),
                Vec::new(),
            ));
        }
        assert!(raw::put(&mut package, &graph).is_err());
        assert!(load(&package).unwrap().is_none());
    }
}
