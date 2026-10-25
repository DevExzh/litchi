//! Source-checked Custom XML package edits (record 0360).

use crate::signature::refuse_signed_change;
use crate::{Error, Result};
use litchi_opc::{OpcError, OpcPackage, PackURI, TargetMode};
use std::sync::Arc;

use super::codec::{
    is_props_relationship, read_props, require_rel_id, rewrite_props, validate_content_type,
    validate_payload, validate_props, write_props,
};
use super::errors::{not_found, unsupported_target};
use super::model::{Conformance, Item, NewItem, Props, Relationship};
use super::package;
use super::patch::{Commit, Diagnostics, Patch};
use super::snapshot::Snapshot;
use super::validation;

/// A bounded edit over one complete Custom XML package graph.
///
/// The edit shares the source package (an [`OpcPackage`] clone shares every
/// part payload) and never changes it: [`Edit::commit`] publishes a new
/// package, and dropping the edit discards it.
#[derive(Clone)]
pub struct Edit {
    source: OpcPackage,
    before: Snapshot,
    draft: Vec<Item>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("items", &self.draft.len())
            .finish_non_exhaustive()
    }
}

impl Edit {
    /// Capture the package graph and start an isolated edit. O(parts of the
    /// Custom XML graph): the package itself is shared, not copied.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn new(package: &OpcPackage) -> Result<Self> {
        let before = Snapshot::load(package)?;
        Ok(Self {
            source: package.clone(),
            draft: before.items().to_vec(),
            before,
        })
    }

    /// Discard every staged change, keeping the edit.
    pub fn reset(&mut self) {
        self.draft = self.before.items().to_vec();
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
        let source_limits = self.source.read_limits();
        let root = validate_payload(&xml, &source_limits)?;
        let current = self
            .draft
            .get(index)
            .ok_or_else(|| not_found("custom_xml.set_xml"))?;
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
            .ok_or_else(|| not_found("custom_xml.set_content_type"))?;
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
        let source_limits = self.source.read_limits();
        let props = read_props(&xml, &source_limits)?;
        let current = self
            .draft
            .get(index)
            .ok_or_else(|| not_found("custom_xml.set_properties_xml"))?;
        if current.props_part().is_none() {
            return Err(unsupported_target(
                "custom_xml.set_properties_xml",
                "the Custom XML item has no properties target",
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
    /// and unknown markup must be authored explicitly. Schema-reference edits
    /// refuse an opaque source part that cannot be rewritten losslessly.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn set_properties(&mut self, index: usize, props: Props) -> Result<bool> {
        validate_props(&props)?;
        let current = self
            .draft
            .get(index)
            .ok_or_else(|| not_found("custom_xml.set_properties"))?;
        let before = current.props().ok_or_else(|| {
            unsupported_target(
                "custom_xml.set_properties",
                "the Custom XML item has no typed properties target",
            )
        })?;
        let source = current.props_xml().ok_or_else(|| {
            unsupported_target(
                "custom_xml.set_properties",
                "the Custom XML item has no properties source",
            )
        })?;
        let conformance = if source
            .windows(super::model::STRICT_NAMESPACE.len())
            .any(|window| window == super::model::STRICT_NAMESPACE.as_bytes())
        {
            Conformance::Strict
        } else {
            Conformance::Transitional
        };
        let source_limits = self.source.read_limits();
        let xml = rewrite_props(source, before, &props, conformance, &source_limits)?;
        drop(props);
        self.set_properties_xml(index, xml)
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
            .ok_or_else(|| not_found("custom_xml.remove_properties"))?;
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
        let item = new_item(&self.source, value)?;
        if self.draft.iter().any(|existing| {
            existing.source() == item.source() && existing.rel_id() == item.rel_id()
        }) {
            return Err(unsupported_target(
                "custom_xml.insert",
                "the Custom XML relationship target is already occupied",
            ));
        }
        if self.draft.iter().any(|existing| {
            existing.part() == item.part()
                || item.props_part().is_some_and(|part| {
                    existing.part() == part || existing.props_part() == Some(part)
                })
        }) {
            return Err(unsupported_target(
                "custom_xml.insert",
                "the Custom XML data target is already occupied",
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
            self.before = Snapshot::load_scoped(&self.source, &hints)?;
        }
        self.draft.push(item);
        let source_limits = self.source.read_limits();
        validation::items(&self.draft, &source_limits)?;
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

    fn is_staged(&self) -> bool {
        self.before.items() != self.draft.as_slice()
    }

    /// Validate the staged graph and publish it as a new package.
    ///
    /// An unchanged edit publishes the source package itself, so a signed
    /// package keeps its signature; a changed one is refused for a signed
    /// package (ADR 0006): remove the signature explicitly first, with
    /// [`OpcPackage::unsign`] or the format facade's `unsign`.
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a change to a signed package, and an
    /// error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_staged() {
            let patch = Patch::new(self.before.clone(), self.before);
            return Ok(Commit::new(self.source, patch, Diagnostics::UNCHANGED));
        }
        let source_limits = self.source.read_limits();
        validation::items(&self.draft, &source_limits)?;
        refuse_signed_change(&self.source)?;

        let mut candidate = self.source;
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
            return Err(Error::Internal(
                "custom XML publication changed the staged occurrence set",
            ));
        }
        let patch = Patch::new(self.before, after);
        Ok(Commit::new(candidate, patch, Diagnostics::CHANGED))
    }
}

fn new_item(package: &OpcPackage, value: NewItem) -> Result<Item> {
    let source_limits = package.read_limits();
    validate_content_type(&value.content_type)?;
    let root = validate_payload(&value.xml, &source_limits)?;
    require_rel_id(&value.rel_id, "custom XML relationship")?;
    match package.get_part(&value.source) {
        Ok(_) => {},
        Err(OpcError::PartNotFound(_)) => {
            return Err(not_found("custom_xml.insert"));
        },
        Err(error) => return Err(error.into()),
    }
    if value.part.as_str() == "/" {
        return Err(unsupported_target(
            "custom_xml.insert",
            "the Custom XML data target cannot be the package root",
        ));
    }
    validate_new_part_name(package, &value.part, "custom_xml.insert")?;

    let (props_part, props, props_xml, mut relationships) = if let Some(new_props) = value.props {
        validate_props(&new_props.value)?;
        require_rel_id(&new_props.rel_id, "custom XML properties relationship")?;
        validate_new_part_name(package, &new_props.part, "custom_xml.insert")?;
        let props_xml = write_props(&new_props.value, value.conformance, &source_limits)?;
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

fn validate_new_part_name(
    package: &OpcPackage,
    part: &PackURI,
    operation: &'static str,
) -> Result<()> {
    match package.validate_new_part_name(part) {
        Ok(()) => Ok(()),
        Err(
            OpcError::DuplicatePartName(_)
            | OpcError::EquivalentPartNames { .. }
            | OpcError::DerivedPartNames { .. },
        ) => Err(unsupported_target(
            operation,
            "the Custom XML destination part is already occupied",
        )),
        Err(error) => Err(error.into()),
    }
}

fn push_name(names: &mut Vec<PackURI>, name: &PackURI) {
    if !names
        .iter()
        .any(|candidate| candidate.as_str().eq_ignore_ascii_case(name.as_str()))
    {
        names.push(name.clone());
    }
}
