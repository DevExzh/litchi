//! Source-checked Custom XML package transactions.

use crate::{Error, Result};
use litchi_opc::{OpcPackage, PackURI, TargetMode};
use std::sync::Arc;

use super::codec::{
    is_props_relationship, read_props, require_rel_id, rewrite_props, validate_content_type,
    validate_payload, validate_props, write_props,
};
use super::model::{Conformance, Item, NewItem, Props, Relationship};
use super::package;
use super::patch::{Commit, Patch};
use super::snapshot::Snapshot;
use super::validation;

/// A bounded edit over one complete Custom XML package graph.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    draft: Vec<Item>,
}

impl<'a> Transaction<'a> {
    /// Capture the package graph and start an isolated transaction.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn new(target: &'a mut OpcPackage) -> Result<Self> {
        let before = Snapshot::load(target)?;
        Ok(Self {
            target,
            draft: before.items().to_vec(),
            before,
        })
    }

    /// Immutable source snapshot used for conflict checks and inverse patches.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the currently staged items.
    #[must_use]
    pub fn items(&self) -> &[Item] {
        &self.draft
    }

    /// Replace one inert data-part XML payload after bounded validation.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn set_xml(&mut self, index: usize, xml: Vec<u8>) -> Result<bool> {
        let root = validate_payload(&xml)?;
        let current = self
            .draft
            .get(index)
            .ok_or_else(|| Error::Invalid(format!("custom XML item index {index} is absent")))?;
        if current.xml() == xml.as_slice() {
            return Ok(false);
        }
        let replacement = clone_item(
            current,
            current.content_type().into(),
            xml,
            root,
            current.props().cloned(),
            current.props_xml().map(ToOwned::to_owned),
            current.relationships().to_vec(),
        );
        self.draft[index] = replacement;
        Ok(true)
    }

    /// Contextual alias for [`Self::set_xml`].
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn set_item_xml(&mut self, index: usize, xml: Vec<u8>) -> Result<bool> {
        self.set_xml(index, xml)
    }

    /// Replace the declared XML content type while retaining payload bytes.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn set_content_type(&mut self, index: usize, content_type: String) -> Result<bool> {
        validate_content_type(&content_type)?;
        let current = self
            .draft
            .get(index)
            .ok_or_else(|| Error::Invalid(format!("custom XML item index {index} is absent")))?;
        if current.content_type() == content_type {
            return Ok(false);
        }
        let replacement = clone_item(
            current,
            content_type,
            current.xml().to_vec(),
            current.root().clone(),
            current.props().cloned(),
            current.props_xml().map(ToOwned::to_owned),
            current.relationships().to_vec(),
        );
        self.draft[index] = replacement;
        Ok(true)
    }

    /// Replace an existing properties payload with caller-provided XML.
    ///
    /// The typed `Props` projection is checked, but every unknown attribute,
    /// child, prefix, comment, and whitespace in the supplied payload remains
    /// byte-for-byte intact.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn set_properties_xml(&mut self, index: usize, xml: Vec<u8>) -> Result<bool> {
        let props = read_props(&xml)?;
        let current = self
            .draft
            .get(index)
            .ok_or_else(|| Error::Invalid(format!("custom XML item index {index} is absent")))?;
        if current.props_part().is_none() {
            return Err(Error::Invalid(
                "custom XML item has no properties relationship".into(),
            ));
        }
        if current.props_xml() == Some(xml.as_slice()) {
            return Ok(false);
        }
        let replacement = clone_item(
            current,
            current.content_type().into(),
            current.xml().to_vec(),
            current.root().clone(),
            Some(props),
            Some(xml),
            current.relationships().to_vec(),
        );
        self.draft[index] = replacement;
        Ok(true)
    }

    /// Replace the typed properties projection using the source conformance
    /// family. Use [`Self::set_properties_xml`] when custom XML lexical form
    /// and unknown markup must be authored explicitly.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn set_properties(&mut self, index: usize, props: Props) -> Result<bool> {
        validate_props(&props)?;
        let current = self
            .draft
            .get(index)
            .ok_or_else(|| Error::Invalid(format!("custom XML item index {index} is absent")))?;
        let before = current.props().ok_or_else(|| {
            Error::Invalid("custom XML item has no typed properties projection".into())
        })?;
        let source = current
            .props_xml()
            .ok_or_else(|| Error::Invalid("custom XML item has no properties source XML".into()))?;
        let conformance = if source
            .windows(super::model::STRICT_NAMESPACE.len())
            .any(|window| window == super::model::STRICT_NAMESPACE.as_bytes())
        {
            Conformance::Strict
        } else {
            Conformance::Transitional
        };
        let xml = rewrite_props(source, before, &props, conformance)?;
        drop(props);
        self.set_properties_xml(index, xml)
    }

    /// Contextual alias for [`Self::set_properties`].
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn set_props(&mut self, index: usize, props: Props) -> Result<bool> {
        self.set_properties(index, props)
    }

    /// Remove the known properties relationship and leave unknown data-part
    /// relationships untouched.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn remove_properties(&mut self, index: usize) -> Result<bool> {
        let current = self
            .draft
            .get(index)
            .ok_or_else(|| Error::Invalid(format!("custom XML item index {index} is absent")))?;
        if current.props_part().is_none() {
            return Ok(false);
        }
        let relationships = current
            .relationships()
            .iter()
            .filter(|relationship| !is_props_relationship(&relationship.relationship_type))
            .cloned()
            .collect();
        self.draft[index] = clone_item(
            current,
            current.content_type().into(),
            current.xml().to_vec(),
            current.root().clone(),
            None,
            None,
            relationships,
        );
        Ok(true)
    }

    /// Add a new data relationship and its optional properties graph.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn insert(&mut self, value: NewItem) -> Result<usize> {
        let item = new_item(self.target, value)?;
        if self.draft.iter().any(|existing| {
            existing.source() == item.source() && existing.rel_id() == item.rel_id()
        }) {
            return Err(Error::Relationship(format!(
                "custom XML relationship '{}' already exists on '{}'",
                item.rel_id(),
                item.source().as_str()
            )));
        }
        if self.draft.iter().any(|existing| {
            existing.part() == item.part()
                || item.props_part().is_some_and(|part| {
                    existing.part() == part || existing.props_part() == Some(part)
                })
        }) {
            return Err(Error::Invalid(
                "custom XML insertion reuses an existing data graph part".into(),
            ));
        }
        if !self
            .before
            .scope_names()
            .iter()
            .any(|name| name == item.source())
        {
            let mut hints = self.before.scope_names();
            push_name(&mut hints, item.source());
            self.before = Snapshot::load_scoped(self.target, &hints)?;
        }
        self.draft.push(item);
        validation::items(&self.draft)?;
        Ok(self.draft.len() - 1)
    }

    /// Remove one data relationship and clean up unreferenced owned parts at
    /// publication time.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn remove(&mut self, index: usize) -> Result<Option<Item>> {
        if index >= self.draft.len() {
            return Ok(None);
        }
        Ok(Some(self.draft.remove(index)))
    }

    /// Whether the staged occurrence set differs from the source snapshot.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before.items() != self.draft.as_slice()
    }

    /// Validate and atomically publish the transaction.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn commit(self) -> Result<Commit> {
        let current = Snapshot::load_scoped(self.target, &self.before.scope_names())?;
        if !self.before.same_source(&current) {
            return Err(super::snapshot::source_mismatch());
        }
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }
        validation::items(&self.draft)?;

        let mut candidate = self.target.clone();
        package::apply_items(&mut candidate, &self.before, &self.draft)?;
        let mut hints = self.before.scope_names();
        for item in &self.draft {
            push_name(&mut hints, item.source());
            push_name(&mut hints, item.part());
            if let Some(props_part) = item.props_part() {
                push_name(&mut hints, props_part);
            }
        }
        let after = Snapshot::load_scoped(&candidate, &hints)?;
        if after.items() != self.draft.as_slice() {
            return Err(Error::Invalid(
                "custom XML publication changed the staged occurrence set".into(),
            ));
        }
        let patch = Patch::new(self.before, after.clone());
        *self.target = candidate;
        Ok(Commit::new(after, patch, true))
    }
}

fn new_item(package: &OpcPackage, value: NewItem) -> Result<Item> {
    validate_content_type(&value.content_type)?;
    let root = validate_payload(&value.xml)?;
    require_rel_id(&value.rel_id, "custom XML relationship")?;
    if package.get_part(&value.source).is_err() {
        return Err(Error::Missing(format!(
            "custom XML source '{}' is absent",
            value.source.as_str()
        )));
    }
    if value.part.as_str() == "/" {
        return Err(Error::Invalid("custom XML data part cannot be root".into()));
    }
    package.validate_new_part_name(&value.part)?;

    let (props_part, props, props_xml, mut relationships) = if let Some(new_props) = value.props {
        validate_props(&new_props.value)?;
        require_rel_id(&new_props.rel_id, "custom XML properties relationship")?;
        package.validate_new_part_name(&new_props.part)?;
        let props_xml = write_props(&new_props.value, value.conformance)?;
        let relationship = Relationship {
            id: new_props.rel_id,
            relationship_type: value.conformance.props_relationship().into(),
            target: new_props.part.relative_ref(value.part.base_uri()),
            target_mode: TargetMode::Internal,
        };
        (
            Some(new_props.part),
            Some(new_props.value),
            Some(props_xml),
            vec![relationship],
        )
    } else {
        (None, None, None, Vec::new())
    };
    relationships.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    let source_relationship = Relationship {
        id: value.rel_id.clone(),
        relationship_type: value.conformance.relationship().into(),
        target: value.part.relative_ref(value.source.base_uri()),
        target_mode: TargetMode::Internal,
    };
    Ok(Item::new(
        value.source,
        value.rel_id,
        source_relationship,
        value.part,
        value.content_type,
        root,
        Arc::new(value.xml),
        props_part,
        props,
        props_xml.map(Arc::new),
        relationships.into(),
    ))
}

fn clone_item(
    current: &Item,
    content_type: String,
    xml: Vec<u8>,
    root: crate::mce::Name,
    props: Option<Props>,
    props_xml: Option<Vec<u8>>,
    relationships: Vec<Relationship>,
) -> Item {
    Item::new(
        current.source().clone(),
        current.rel_id().into(),
        current.source_relationship().clone(),
        current.part().clone(),
        content_type,
        root,
        Arc::new(xml),
        current.props_part().cloned(),
        props,
        props_xml.map(Arc::new),
        relationships.into(),
    )
}

fn push_name(names: &mut Vec<PackURI>, name: &PackURI) {
    if !names
        .iter()
        .any(|candidate| candidate.as_str().eq_ignore_ascii_case(name.as_str()))
    {
        names.push(name.clone());
    }
}
