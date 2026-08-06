//! Immutable source-bound Custom XML package snapshots.

use crate::{Error, Result};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part};
use std::sync::Arc;

use super::model::{Item, Relationship};
use super::{package, validation};

/// An immutable semantic and physical snapshot of the Custom XML graph.
///
/// XML payloads and every captured relationship target remain opaque. Cloning
/// a snapshot shares part bytes through `Arc` and copies only bounded graph
/// metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    items: Vec<Item>,
    source: SourceState,
}

impl Snapshot {
    /// Discover and validate the package's explicit Custom XML graph.
    pub fn load(package: &OpcPackage) -> Result<Self> {
        Self::load_scoped(package, &[])
    }

    /// Alias for [`Self::load`] emphasizing source-bound parsing.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        Self::load(package)
    }

    /// Borrow items in stable source-part and relationship-ID order.
    #[must_use]
    pub fn items(&self) -> &[Item] {
        &self.items
    }

    /// Look up one item by source order.
    #[must_use]
    pub fn item(&self, index: usize) -> Option<&Item> {
        self.items.get(index)
    }

    /// Find an item by its source relationship ID.
    #[must_use]
    pub fn by_relationship(&self, source: &PackURI, rel_id: &str) -> Option<&Item> {
        self.items
            .iter()
            .find(|item| item.source() == source && item.rel_id() == rel_id)
    }

    /// Find the first item whose typed properties carry `itemID`.
    #[must_use]
    pub fn by_item_id(&self, id: &str) -> Option<&Item> {
        self.items.iter().find(|item| {
            item.props()
                .is_some_and(|props| props.id.eq_ignore_ascii_case(id))
        })
    }

    /// Whether no explicit Custom XML relationships were discovered.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.items.is_empty()
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.items == other.items && self.source == other.source
    }

    pub(crate) fn scope_names(&self) -> Vec<PackURI> {
        self.source
            .parts
            .iter()
            .map(|part| part.name.clone())
            .collect()
    }

    pub(crate) fn source(&self) -> &SourceState {
        &self.source
    }

    pub(crate) fn load_scoped(package: &OpcPackage, hints: &[PackURI]) -> Result<Self> {
        let items = package::discover(package)?;
        validation::items(&items)?;

        let mut names = Vec::with_capacity(hints.len() + items.len() * 3);
        for name in hints {
            push_unique(&mut names, name);
        }
        for item in &items {
            push_unique(&mut names, item.source());
            push_unique(&mut names, item.part());
            if let Some(props_part) = item.props_part() {
                push_unique(&mut names, props_part);
            }
        }
        names.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));

        let mut parts = Vec::with_capacity(names.len());
        for name in names {
            let Ok(part) = package.get_part(&name) else {
                // Hints intentionally include parts that a remove operation
                // may have detached. Discovered item-owned parts are checked
                // by `discover` above and therefore cannot be missing here.
                continue;
            };
            parts.push(PartState::capture(part));
        }
        let source = SourceState { parts };
        Ok(Self { items, source })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct SourceState {
    pub(crate) parts: Vec<PartState>,
}

impl SourceState {
    pub(crate) fn part(&self, name: &PackURI) -> Option<&PartState> {
        self.parts.iter().find(|part| part.name == *name)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct PartState {
    pub(crate) name: PackURI,
    pub(crate) content_type: String,
    pub(crate) data: Arc<Vec<u8>>,
    pub(crate) relationships: Vec<Relationship>,
}

impl PartState {
    pub(crate) fn capture(part: &dyn Part) -> Self {
        let mut relationships = part
            .rels()
            .iter()
            .map(Relationship::from_opc)
            .collect::<Vec<_>>();
        relationships.sort_unstable_by(|left, right| left.id.cmp(&right.id));
        Self {
            name: part.partname().clone(),
            content_type: part.content_type().into(),
            data: part.blob_arc(),
            relationships,
        }
    }

    pub(crate) fn to_part(&self) -> Result<BlobPart> {
        let mut part = BlobPart::new_shared(
            self.name.clone(),
            self.content_type.clone(),
            Arc::clone(&self.data),
        );
        for relationship in &self.relationships {
            part.rels_mut().try_add_relationship(
                relationship.relationship_type.clone(),
                relationship.target.clone(),
                relationship.id.clone(),
                relationship.target_mode,
            )?;
        }
        Ok(part)
    }

    pub(crate) fn matches(&self, part: &dyn Part) -> bool {
        self == &Self::capture(part)
    }
}

fn push_unique(names: &mut Vec<PackURI>, candidate: &PackURI) {
    if !names
        .iter()
        .any(|name| name.as_str().eq_ignore_ascii_case(candidate.as_str()))
    {
        names.push(candidate.clone());
    }
}

pub(crate) fn source_mismatch() -> Error {
    Error::Relationship("custom XML patch source graph changed".into())
}
