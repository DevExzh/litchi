#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_same,
    reason = "the validated binding intentionally replaces its fallible precursor"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
//! Typed semantic glossary catalog and failure-atomic in-memory CRUD.

use super::codec::{
    GALLERIES, add_sizes, authored_analysis, background_analysis, bounded, canonical_id,
    entry_analysis, invalid, name_key, replace_sizes, validate_authored_name,
    validate_catalog_sizes, validate_name,
};
use super::graph::raw;
use super::{Arc, Error, HashMap, MAX_PARTS, R, REL, RS, Result, STRICT_REL, W, WS};
/// Namespace and relationship family used by a glossary part.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub(in crate::glossary) fn index(self) -> usize {
        match self {
            Self::Transitional => 0,
            Self::Strict => 1,
        }
    }

    pub(in crate::glossary) fn word(self) -> &'static str {
        match self {
            Self::Transitional => W,
            Self::Strict => WS,
        }
    }

    pub(in crate::glossary) fn relationships(self) -> &'static str {
        match self {
            Self::Transitional => R,
            Self::Strict => RS,
        }
    }

    pub(in crate::glossary) fn glossary_relationship(self) -> &'static str {
        match self {
            Self::Transitional => REL,
            Self::Strict => STRICT_REL,
        }
    }

    pub(in crate::glossary) fn from_word(namespace: &str) -> Result<Self> {
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if GALLERIES.contains(&value.as_str()) {
            Ok(Self(value))
        } else {
            Err(invalid(format!("invalid document-part gallery '{value}'")))
        }
    }

    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Name {
    pub(in crate::glossary) value: String,
    pub(in crate::glossary) key: Arc<str>,
    pub(in crate::glossary) decorated: Option<bool>,
}
impl Name {
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
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

    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.value
    }

    #[must_use]
    pub fn decorated(&self) -> Option<bool> {
        self.decorated
    }

    fn key(&self) -> &str {
        &self.key
    }

    #[must_use]
    pub fn with_decorated(mut self, decorated: bool) -> Self {
        self.decorated = Some(decorated);
        self
    }
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Category {
    pub(in crate::glossary) name: String,
    pub(in crate::glossary) gallery: Gallery,
}
impl Category {
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn new(name: impl Into<String>, gallery: Gallery) -> Result<Self> {
        let name = name.into();
        bounded(&name)?;
        if name.trim().is_empty() {
            return Err(invalid("building-block category name cannot be empty"));
        }
        Ok(Self { name, gallery })
    }

    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    #[must_use]
    pub fn gallery(&self) -> &Gallery {
        &self.gallery
    }
}

/// Canonical braced uppercase building-block GUID.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct Id(String);

impl Id {
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
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

    #[must_use]
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
    #[must_use]
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

    #[must_use]
    pub fn name(&self) -> Option<&Name> {
        self.name.as_ref()
    }

    #[must_use]
    pub fn with_name(mut self, name: Name) -> Self {
        self.name = Some(name);
        self
    }
}
#[derive(Clone, Debug)]
pub struct Entry {
    pub(in crate::glossary) props: Option<Props>,
    /// Full inert `w:docPartBody` subtree.
    pub(in crate::glossary) body: Option<Vec<u8>>,
    pub(in crate::glossary) producer: Option<Box<ProducerEntry>>,
    pub(in crate::glossary) sizes: [usize; 2],
    pub(in crate::glossary) refs: Arc<[String]>,
    pub(in crate::glossary) lineage: Option<Arc<Lineage>>,
}

#[derive(Clone, Debug)]
pub(in crate::glossary) struct ProducerEntry {
    pub(in crate::glossary) conformance: Conformance,
    pub(in crate::glossary) xml: Arc<str>,
    pub(in crate::glossary) refs: Arc<[String]>,
}

#[derive(Debug)]
pub(in crate::glossary) struct Lineage;

pub(in crate::glossary) struct EntryAnalysis {
    pub(in crate::glossary) sizes: [usize; 2],
    pub(in crate::glossary) refs: Arc<[String]>,
}
#[derive(Clone, Debug)]
pub struct Catalog {
    /// Full inert `w:background` subtree.
    pub(in crate::glossary) background: Option<Vec<u8>>,
    pub(in crate::glossary) background_refs: Arc<[String]>,
    pub(in crate::glossary) background_lineage: Option<Arc<Lineage>>,
    pub(in crate::glossary) entries: Vec<Entry>,
    /// Physical resource binding captured by package-aware loads.
    pub(in crate::glossary) binding: Option<Box<Binding>>,
    pub(in crate::glossary) state: CatalogState,
}

#[derive(Clone, Debug, Default)]
pub(in crate::glossary) struct CatalogState {
    pub(in crate::glossary) names: HashMap<Arc<str>, NameMatch>,
    pub(in crate::glossary) entry_bytes: [usize; 2],
    pub(in crate::glossary) background_bytes: [usize; 2],
}

#[derive(Clone, Copy, Debug)]
pub(in crate::glossary) struct NameMatch {
    pub(in crate::glossary) first: usize,
    pub(in crate::glossary) count: usize,
}

#[derive(Clone, Debug)]
pub(in crate::glossary) struct Binding {
    pub(in crate::glossary) conformance: Conformance,
    pub(in crate::glossary) rels: Vec<raw::Rel>,
    pub(in crate::glossary) parts: Vec<raw::Part>,
    pub(in crate::glossary) root_name: String,
    pub(in crate::glossary) owner_main: Option<String>,
    pub(in crate::glossary) owner_id: Option<String>,
    pub(in crate::glossary) owner_target: Option<String>,
    pub(in crate::glossary) lineage: Arc<Lineage>,
}

impl Binding {
    pub(in crate::glossary) fn from_graph(graph: &raw::Graph) -> Self {
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

    pub(in crate::glossary) fn matches(&self, graph: &raw::Graph) -> bool {
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
    pub(in crate::glossary) fn rebuild_state(&mut self) -> Result<()> {
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

    #[must_use]
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

    #[must_use]
    pub fn background(&self) -> Option<&[u8]> {
        self.background.as_deref()
    }

    pub(in crate::glossary) fn validate_entry_lineage(&self, entry: &Entry) -> Result<()> {
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

    pub(in crate::glossary) fn validate_bound_lineages(&self) -> Result<()> {
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

    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
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

    #[must_use]
    pub fn entries(&self) -> &[Entry] {
        &self.entries
    }

    #[must_use]
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Entry> {
        self.entries.iter()
    }

    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Find a uniquely named entry using Unicode default case folding.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn get(&self, name: &str) -> Result<Option<&Entry>> {
        Ok(self
            .state
            .offset(name)?
            .and_then(|offset| self.entries.get(offset)))
    }

    /// Checked numeric fallback for import and inspection workflows.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn at(&self, index: usize) -> Result<&Entry> {
        self.entries.get(index).ok_or(Error::OutOfBounds {
            object: "glossary entry",
            index,
            len: self.entries.len(),
        })
    }

    /// Add a fresh, uniquely named entry by moving it into the catalog.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn replace(&mut self, name: &str, entry: Entry) -> Result<Option<Entry>> {
        let Some(index) = self.state.offset(name)? else {
            return Ok(None);
        };
        self.replace_at(index, entry).map(Some)
    }

    /// Checked numeric fallback for replacement and ambiguous-name repair.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn rename(&mut self, name: &str, replacement: Name) -> Result<bool> {
        let Some(index) = self.state.offset(name)? else {
            return Ok(false);
        };
        self.rename_at(index, replacement)
    }

    /// Checked numeric fallback for renaming an ambiguous producer entry.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
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

    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn remove(&mut self, name: &str) -> Result<Option<Entry>> {
        let Some(offset) = self.state.offset(name)? else {
            return Ok(None);
        };
        self.remove_at(offset).map(Some)
    }

    /// Checked numeric fallback for removal.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
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

    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn move_to(&mut self, name: &str, to: usize) -> Result<bool> {
        let Some(from) = self.state.offset(name)? else {
            return Ok(false);
        };
        self.move_at(from, to)?;
        Ok(true)
    }

    /// Checked numeric fallback for reordering.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
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

pub(in crate::glossary) fn entry_key(entry: &Entry) -> Option<&str> {
    entry
        .props
        .as_ref()
        .and_then(|props| props.name.as_ref())
        .map(Name::key)
}

pub(in crate::glossary) fn restore_name(
    entry: &mut Entry,
    name: Name,
    producer: Option<Box<ProducerEntry>>,
) -> Result<()> {
    let props = entry
        .props
        .as_mut()
        .ok_or_else(|| invalid("building block properties disappeared during rename rollback"))?;
    props.name = Some(name);
    entry.producer = producer;
    Ok(())
}

impl Entry {
    pub(in crate::glossary) fn has_relationship_references(&self) -> bool {
        !self.refs.is_empty()
            || self
                .producer
                .as_deref()
                .is_some_and(|producer| !producer.refs.is_empty())
    }

    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
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

    #[must_use]
    pub fn props(&self) -> Option<&Props> {
        self.props.as_ref()
    }

    #[must_use]
    pub fn body(&self) -> Option<&[u8]> {
        self.body.as_deref()
    }

    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
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

    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
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

    #[must_use]
    pub fn into_body(self) -> Option<Vec<u8>> {
        self.body
    }
}
