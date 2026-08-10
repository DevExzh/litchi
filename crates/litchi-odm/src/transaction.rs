//! Unified title and linked-section transactions.

use litchi_core::{
    BlobBundle, BlobId, BlobLimits, ConflictSet, Error, ForwardOnly, History as CoreHistory,
    HistoryLimits, Patch as CorePatch, PatchLimits, PatchOperation, Position, Result, Reversible,
    ReversibleOperation,
};
use serde_json::Value;
use std::{
    collections::{BTreeMap, HashMap, HashSet},
    sync::Arc,
};

use crate::{Master, link::Selector};

pub use crate::edit_ops::{
    ActiveContentPolicy, BodyItemChange, BodyItemProvenance, BodyItemSpec, GeneratedIndexChange,
    ResourceChange, ResourceSpec, SectionChange, SectionSpec, SecurityPolicy, StyleChange,
    StyleSpec, SubdocumentSpec,
};

const MAX_PACKAGE_BYTES: usize = 256 * 1024 * 1024;
const MAX_WIRE_JSON_BYTES: usize = 768 * 1024 * 1024;
const DURABLE_FORMAT: &str = "litchi.odm";
const DURABLE_OPERATION: &str = "package.replace";
const DURABLE_TARGET: &str = "master";
const SOURCE_PRECONDITION: &str = "source_sha256";

/// How a cross-master transfer handles an occupied identity.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum CollisionPolicy {
    /// Refuse every occupied destination identity.
    Refuse,
    /// Reuse an identity only when its complete bytes/definition match.
    ReuseIdentical,
    /// Reuse identical content or choose a deterministic imported identity.
    Rename,
}

/// Collision behavior for a dependency-closed linked-section transfer.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TransferOptions {
    resource: CollisionPolicy,
    style: CollisionPolicy,
}

impl TransferOptions {
    /// Uses deterministic rename with identical-content reuse.
    #[must_use]
    pub const fn collision_safe() -> Self {
        Self {
            resource: CollisionPolicy::Rename,
            style: CollisionPolicy::Rename,
        }
    }

    /// Configures package-resource collisions.
    #[must_use]
    pub const fn with_resource_collision(mut self, policy: CollisionPolicy) -> Self {
        self.resource = policy;
        self
    }

    /// Configures style-identity collisions.
    #[must_use]
    pub const fn with_style_collision(mut self, policy: CollisionPolicy) -> Self {
        self.style = policy;
        self
    }
}

impl Default for TransferOptions {
    fn default() -> Self {
        Self::collision_safe()
    }
}

/// One atomic edit derived from an immutable master snapshot.
pub struct Edit<'source> {
    source: &'source Master,
    title_before: Option<String>,
    title_after: Option<String>,
    links: BTreeMap<usize, String>,
    metadata: Option<litchi_core::Metadata>,
    sections: Vec<SectionChange>,
    generated_indexes: Vec<GeneratedIndexChange>,
    body_items: Vec<BodyItemChange>,
    styles: Vec<StyleChange>,
    resources: BTreeMap<String, ResourceChange>,
    policy: SecurityPolicy,
}

impl<'source> Edit<'source> {
    pub(crate) fn new(source: &'source Master) -> Self {
        Self::with_policy(source, SecurityPolicy::default())
    }

    pub(crate) fn with_policy(source: &'source Master, policy: SecurityPolicy) -> Self {
        let title = source.title().map(str::to_owned);
        Self {
            source,
            title_before: title.clone(),
            title_after: title,
            links: BTreeMap::new(),
            metadata: None,
            sections: Vec::new(),
            generated_indexes: Vec::new(),
            body_items: Vec::new(),
            styles: Vec::new(),
            resources: BTreeMap::new(),
            policy,
        }
    }

    /// Returns the staged title.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title_after.as_deref()
    }

    /// Stages a bounded XML 1.0 title.
    ///
    /// # Errors
    ///
    /// Returns an error when the title exceeds the limit or contains a
    /// character forbidden by XML 1.0.
    pub fn set_title(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        let title = value.into();
        crate::title::validate_title(&title)?;
        self.title_after = Some(title.clone());
        if let Some(metadata) = &mut self.metadata {
            metadata.title = Some(title);
        }
        Ok(self)
    }

    /// Stages removal of the title element.
    pub fn clear_title(&mut self) -> &mut Self {
        self.title_after = None;
        if let Some(metadata) = &mut self.metadata {
            metadata.title = None;
        }
        self
    }

    /// Stages one existing linked-section target by exact section name or
    /// checked semantic position.
    ///
    /// Targets remain inert and are never resolved, opened, or fetched.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing selector or invalid XML target.
    pub fn set_link<'selector>(
        &mut self,
        selector: impl Into<Selector<'selector>>,
        href: impl Into<String>,
    ) -> Result<&mut Self> {
        let reference = resolve(self.source, selector.into())?;
        let href = href.into();
        crate::link::validate_href(&href)?;
        if !self.policy.allows_external_targets() && is_external_target(&href) {
            return Err(invalid(
                "ODM security policy refuses an external link target",
            ));
        }
        self.links.insert(reference.get(), href);
        Ok(self)
    }

    /// Stages the five mutable simple metadata fields: title, author, subject,
    /// description, and keywords. Other fields remain source-preserved.
    ///
    /// # Errors
    ///
    /// Returns an error when a staged text field violates XML bounds.
    pub fn set_metadata(&mut self, metadata: litchi_core::Metadata) -> Result<&mut Self> {
        for (value, scope) in [
            (metadata.title.as_deref(), "ODM metadata title"),
            (metadata.author.as_deref(), "ODM metadata author"),
            (metadata.subject.as_deref(), "ODM metadata subject"),
            (metadata.description.as_deref(), "ODM metadata description"),
            (metadata.keywords.as_deref(), "ODM metadata keywords"),
        ] {
            if let Some(value) = value {
                crate::edit_ops::validate_value(value, scope, true)?;
            }
        }
        self.title_after.clone_from(&metadata.title);
        self.metadata = Some(metadata);
        Ok(self)
    }

    /// Stages insertion of an empty root section.
    ///
    /// # Errors
    ///
    /// Returns an error for a duplicate section name or missing style.
    pub fn add_section(&mut self, section: SectionSpec) -> Result<&mut Self> {
        if section_name_exists(self.source, &self.sections, section.name()) {
            return Err(invalid("ODM section destination name already exists"));
        }
        if let Some(style_name) = section.style_name()
            && !style_name_exists(self.source, &self.styles, style_name)
        {
            return Err(invalid("ODM section style does not exist"));
        }
        if section
            .subdocument()
            .is_some_and(|subdocument| is_external_target(subdocument.href()))
            && !self.policy.allows_external_targets()
        {
            return Err(invalid(
                "ODM security policy refuses an external linked section",
            ));
        }
        self.sections.push(SectionChange::Add(section));
        Ok(self)
    }

    /// Copies one linked section and its package resource from another master.
    ///
    /// The complete style-parent closure is reused when identical or copied
    /// under deterministic imported names when it differs. Resource
    /// collisions use the same default behavior. The source package is never
    /// recursively opened or executed.
    ///
    /// # Errors
    ///
    /// Returns an error for a non-linked or external source section, a missing
    /// style/resource dependency, or a policy violation.
    pub fn transfer_linked_section(
        &mut self,
        source: &Master,
        position: Position,
        destination_name: impl Into<String>,
        destination_path: impl Into<String>,
    ) -> Result<&mut Self> {
        self.transfer_linked_section_with_options(
            source,
            position,
            destination_name,
            destination_path,
            TransferOptions::default(),
        )
    }

    /// Copies a linked section, its complete style-parent closure, and its
    /// package resource under explicit collision rules.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid source semantics, a refused collision, or
    /// a transaction security-policy violation.
    pub fn transfer_linked_section_with_options(
        &mut self,
        source: &Master,
        position: Position,
        destination_name: impl Into<String>,
        destination_path: impl Into<String>,
        options: TransferOptions,
    ) -> Result<&mut Self> {
        let destination_name = destination_name.into();
        SectionSpec::new(&destination_name)?;
        if section_name_exists(self.source, &self.sections, &destination_name) {
            return Err(invalid("ODM section destination name already exists"));
        }
        let node = source
            .section_tree()
            .get(position)
            .ok_or_else(|| invalid("ODM transfer section selector is out of bounds"))?;
        let reference_position = node
            .reference()
            .ok_or_else(|| invalid("ODM transfer source section is not linked"))?;
        let reference = source
            .subdocuments()
            .get(reference_position.get())
            .ok_or_else(|| invalid("ODM transfer subdocument dependency disappeared"))?;
        let crate::subdocument::Target::Package(source_path) = reference.target() else {
            return Err(invalid(
                "ODM linked-section transfer requires a package target",
            ));
        };
        let destination_path = destination_path.into();
        let source_resource = source
            .resources()
            .resources()
            .iter()
            .find(|resource| resource.path() == source_path)
            .ok_or_else(|| invalid("ODM transfer package resource is missing"))?;
        let media_type = source_resource
            .media_type()
            .unwrap_or("application/octet-stream");
        let bytes = source.resource_bytes(source_path)?;
        if bytes.len() > self.policy.max_resource_bytes() {
            return Err(invalid(
                "ODM resource exceeds the transaction security policy",
            ));
        }
        let (destination_path, write_resource) = resolve_resource_collision(
            self.source,
            &self.resources,
            &destination_path,
            media_type,
            &bytes,
            options.resource,
        )?;
        let style_name = node
            .style_name()
            .map(|name| transfer_style_closure(self, source, name, options.style))
            .transpose()?;
        if write_resource {
            self.put_resource(ResourceSpec::new(
                destination_path.clone(),
                media_type,
                bytes,
            )?)?;
        }
        let mut subdocument = SubdocumentSpec::new(destination_path)?;
        if let Some(source_section) = reference.source_section() {
            subdocument = subdocument.with_source_section(source_section)?;
        }
        if let Some(filter_name) = reference.filter_name() {
            subdocument = subdocument.with_filter_name(filter_name)?;
        }
        let mut section = SectionSpec::new(destination_name)?.with_subdocument(subdocument);
        if let Some(style_name) = style_name {
            section = section.with_style(style_name)?;
        }
        self.add_section(section)
    }

    /// Renames one source section and its modeled local-section references.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale selector or duplicate destination name.
    pub fn rename_section(
        &mut self,
        position: Position,
        name: impl Into<String>,
    ) -> Result<&mut Self> {
        let name = name.into();
        crate::edit_ops::validate_value(&name, "ODM section name", false)?;
        let before = self
            .source
            .section_tree()
            .get(position)
            .ok_or_else(|| invalid("ODM section selector is out of bounds"))?
            .name()
            .to_owned();
        if before != name && section_name_exists(self.source, &self.sections, &name) {
            return Err(invalid("ODM section destination name already exists"));
        }
        ensure_section_not_staged(&self.sections, position)?;
        self.sections.push(SectionChange::Rename {
            position,
            before,
            after: name,
        });
        Ok(self)
    }

    /// Renames one generated index by its direct master-body item position.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/unnamed index, duplicate destination
    /// name, or invalid bounded XML value.
    pub fn rename_generated_index(
        &mut self,
        item: Position,
        name: impl Into<String>,
    ) -> Result<&mut Self> {
        let index = self
            .source
            .structure()
            .generated_indexes()
            .iter()
            .find(|index| index.item() == item)
            .ok_or_else(|| invalid("ODM generated-index selector was not found"))?;
        let before = index
            .name()
            .ok_or_else(|| invalid("ODM generated index has no text:name"))?
            .to_owned();
        if self.body_items.iter().any(
            |change| matches!(change, BodyItemChange::Remove { item: staged, .. } if *staged == item),
        ) {
            return Err(invalid("ODM generated index is already staged for removal"));
        }
        let after = name.into();
        crate::edit_ops::validate_value(&after, "ODM generated index name", false)?;
        if before != after
            && (self.source.structure().generated_indexes().iter().any(|other| {
                other.item() != item
                    && other
                        .name()
                        .is_some_and(|candidate| candidate == after.as_str())
            }) || self.generated_indexes.iter().any(|change| {
                matches!(change, GeneratedIndexChange::Rename { after: staged, .. } if staged == &after)
            }) || self.body_items.iter().any(|change| {
                matches!(change, BodyItemChange::Add(spec) if spec.generated_index_name() == Some(after.as_str()))
            }))
        {
            return Err(invalid("ODM generated-index destination name already exists"));
        }
        if self.generated_indexes.iter().any(|change| {
            matches!(change, GeneratedIndexChange::Rename { item: staged, .. } if *staged == item)
        }) {
            return Err(invalid("ODM generated index already has a staged change"));
        }
        self.generated_indexes.push(GeneratedIndexChange::Rename {
            item,
            before,
            after,
        });
        Ok(self)
    }

    /// Removes one common direct master-body subtree without reserializing it.
    ///
    /// Paragraphs, headings, lists, tables, generated indexes, and unknown
    /// extension children are removed by their checked source span. Sections
    /// retain their dependency-aware [`Self::remove_section`] operation, while
    /// declaration containers are intentionally not editable through this
    /// generic operation.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale selector, an unsupported item kind, or an
    /// overlapping generated-index change.
    pub fn remove_body_item(&mut self, item: Position) -> Result<&mut Self> {
        let kind = *self
            .source
            .structure()
            .items()
            .get(item.get())
            .ok_or_else(|| invalid("ODM master-body item selector is out of bounds"))?;
        match kind {
            crate::structure::Kind::Section(_) => {
                return Err(invalid(
                    "ODM section removal requires dependency-aware remove_section",
                ));
            },
            crate::structure::Kind::Declarations => {
                return Err(invalid(
                    "ODM declaration containers are not generic body-item edits",
                ));
            },
            crate::structure::Kind::Paragraph
            | crate::structure::Kind::Heading
            | crate::structure::Kind::List
            | crate::structure::Kind::Table
            | crate::structure::Kind::GeneratedIndex(_)
            | crate::structure::Kind::Other => {},
        }
        if self.generated_indexes.iter().any(
            |change| matches!(change, GeneratedIndexChange::Rename { item: staged, .. } if *staged == item),
        ) {
            return Err(invalid("ODM master-body item already has a staged change"));
        }
        if self.body_items.iter().any(
            |change| matches!(change, BodyItemChange::Remove { item: staged, .. } if *staged == item),
        ) {
            return Err(invalid("ODM master-body item already has a staged change"));
        }
        self.body_items.push(BodyItemChange::Remove { item, kind });
        Ok(self)
    }

    /// Appends one typed common item to the direct master body.
    ///
    /// # Errors
    ///
    /// Returns an error for a duplicate generated-index identity.
    pub fn add_body_item(&mut self, item: BodyItemSpec) -> Result<&mut Self> {
        if let Some(name) = item.generated_index_name()
            && generated_index_name_exists(
                self.source,
                &self.generated_indexes,
                &self.body_items,
                name,
            )
        {
            return Err(invalid(
                "ODM generated-index destination name already exists",
            ));
        }
        self.body_items.push(BodyItemChange::Add(item));
        Ok(self)
    }

    /// Copies one dependency-free common body item from an immutable source.
    ///
    /// The exact standalone fragment and SHA-256/item provenance are retained.
    /// Style- or resource-dependent fragments are refused because this narrow
    /// operation does not infer an incomplete dependency closure.
    ///
    /// # Errors
    ///
    /// Returns an error for an unsupported selector, dependency, or identity
    /// collision.
    pub fn transfer_body_item(&mut self, source: &Master, item: Position) -> Result<&mut Self> {
        self.add_body_item(BodyItemSpec::imported(source, item)?)
    }

    /// Removes one section subtree when no modeled local reference targets it.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale selector, overlapping intent, or incoming
    /// local-section dependency.
    pub fn remove_section(&mut self, position: Position) -> Result<&mut Self> {
        let before = self
            .source
            .section_tree()
            .get(position)
            .ok_or_else(|| invalid("ODM section selector is out of bounds"))?
            .name()
            .to_owned();
        ensure_section_not_staged(&self.sections, position)?;
        self.sections
            .push(SectionChange::Remove { position, before });
        Ok(self)
    }

    /// Removes one section subtree and package resources referenced only by it.
    ///
    /// Shared resources remain in place. Local-section dependencies are
    /// checked during atomic publication.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale selector or unreadable resource closure.
    pub fn remove_section_with_orphaned_resources(
        &mut self,
        position: Position,
    ) -> Result<&mut Self> {
        let root = self
            .source
            .section_tree()
            .get(position)
            .ok_or_else(|| invalid("ODM section selector is out of bounds"))?;
        let mut orphaned = Vec::new();
        orphaned
            .try_reserve(self.source.resources().resources().len())
            .map_err(|source| Error::Allocation {
                resource: "ODM orphaned section resources",
                source,
            })?;
        for resource in self.source.resources().resources() {
            if resource.references().is_empty() {
                continue;
            }
            let referenced_only_by_subtree = resource.references().iter().all(|reference| {
                self.source
                    .subdocuments()
                    .get(reference.get())
                    .and_then(|linked| {
                        self.source
                            .section_tree()
                            .sections()
                            .iter()
                            .find(|section| section.name() == linked.section())
                    })
                    .is_some_and(|section| {
                        root.source_span.start <= section.source_span.start
                            && section.source_span.end <= root.source_span.end
                    })
            });
            if referenced_only_by_subtree {
                orphaned.push(resource.path().to_owned());
            }
        }
        for path in orphaned {
            self.remove_resource(path)?;
        }
        self.remove_section(position)
    }

    /// Adds one minimal named style definition.
    ///
    /// # Errors
    ///
    /// Returns an error for a duplicate style identity.
    pub fn add_style(&mut self, style: StyleSpec) -> Result<&mut Self> {
        if style_name_exists(self.source, &self.styles, style.name()) {
            return Err(invalid("ODM style destination name already exists"));
        }
        self.styles.push(StyleChange::Add(style));
        Ok(self)
    }

    /// Renames one style and every modeled style-name/parent-style-name use.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/ambiguous source or duplicate destination.
    pub fn rename_style(
        &mut self,
        origin: crate::style::Origin,
        before: impl Into<String>,
        after: impl Into<String>,
    ) -> Result<&mut Self> {
        let before = before.into();
        let after = after.into();
        crate::edit_ops::validate_value(&after, "ODM style name", false)?;
        resolve_style(self.source, origin, &before)?;
        if before != after && style_name_exists(self.source, &self.styles, &after) {
            return Err(invalid("ODM style destination name already exists"));
        }
        self.styles.push(StyleChange::Rename {
            origin,
            before,
            after,
        });
        Ok(self)
    }

    /// Removes one unreferenced style definition.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/ambiguous style or dependency block.
    pub fn remove_style(
        &mut self,
        origin: crate::style::Origin,
        name: impl Into<String>,
    ) -> Result<&mut Self> {
        let name = name.into();
        resolve_style(self.source, origin, &name)?;
        self.styles.push(StyleChange::Remove { origin, name });
        Ok(self)
    }

    /// Adds or replaces one inert package resource.
    ///
    /// # Errors
    ///
    /// Returns an error when the explicit resource byte policy is exceeded.
    pub fn put_resource(&mut self, resource: ResourceSpec) -> Result<&mut Self> {
        if resource.bytes().len() > self.policy.max_resource_bytes() {
            return Err(invalid(
                "ODM resource exceeds the transaction security policy",
            ));
        }
        self.resources
            .insert(resource.path().to_owned(), ResourceChange::Put(resource));
        Ok(self)
    }

    /// Removes an unreferenced inert package resource.
    ///
    /// # Errors
    ///
    /// Returns an error when the path is absent. Final dependency validation
    /// also refuses a resource still targeted by a linked section.
    pub fn remove_resource(&mut self, path: impl Into<String>) -> Result<&mut Self> {
        let path = path.into();
        if !self
            .source
            .resources()
            .resources()
            .iter()
            .any(|resource| resource.path() == path)
        {
            return Err(invalid("ODM resource path was not found"));
        }
        let resource = self
            .source
            .resources()
            .resources()
            .iter()
            .find(|resource| resource.path() == path)
            .ok_or_else(|| invalid("ODM resource path was not found"))?;
        let previous = ResourceSpec::new(
            path.clone(),
            resource.media_type().unwrap_or("application/octet-stream"),
            self.source.resource_bytes(&path)?,
        )?;
        self.resources
            .insert(path, ResourceChange::Remove(previous));
        Ok(self)
    }

    /// Copies one inert resource from another immutable master snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent source path or policy violation.
    pub fn transfer_resource(
        &mut self,
        source: &Master,
        source_path: &str,
        destination_path: impl Into<String>,
    ) -> Result<&mut Self> {
        self.transfer_resource_with_collision(
            source,
            source_path,
            destination_path,
            CollisionPolicy::Refuse,
        )
    }

    /// Copies one inert resource with explicit collision handling.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent source, refused collision, or policy
    /// violation.
    pub fn transfer_resource_with_collision(
        &mut self,
        source: &Master,
        source_path: &str,
        destination_path: impl Into<String>,
        collision: CollisionPolicy,
    ) -> Result<&mut Self> {
        let resource = source
            .resources()
            .resources()
            .iter()
            .find(|resource| resource.path() == source_path)
            .ok_or_else(|| invalid("ODM transfer source resource was not found"))?;
        let media_type = resource.media_type().unwrap_or("application/octet-stream");
        let bytes = source.resource_bytes(source_path)?;
        let destination_path = destination_path.into();
        let (destination_path, write) = resolve_resource_collision(
            self.source,
            &self.resources,
            &destination_path,
            media_type,
            &bytes,
            collision,
        )?;
        if write {
            self.put_resource(ResourceSpec::new(destination_path, media_type, bytes)?)?;
        }
        Ok(self)
    }

    /// Publishes every staged effect as one fully reopened package.
    ///
    /// # Errors
    ///
    /// Returns an error for signed/encrypted input, invalid staged data, or
    /// semantic readback which differs from the request.
    pub fn commit(self) -> Result<Commit> {
        let title_changed = self.title_before != self.title_after;
        let explicit_metadata = self.metadata.is_some();
        let link_changes = collect_link_changes(self.source, &self.links)?;
        let extended_changed = self.metadata.is_some()
            || !self.sections.is_empty()
            || !self.generated_indexes.is_empty()
            || !self.body_items.is_empty()
            || !self.styles.is_empty()
            || !self.resources.is_empty();
        if !title_changed && link_changes.is_empty() && !extended_changed {
            return Ok(Commit::new(
                self.source,
                self.source.clone(),
                ChangeSet::default(),
            ));
        }

        let requested_metadata = if let Some(metadata) = self.metadata {
            Some(metadata)
        } else if title_changed {
            let mut metadata = self.source.metadata().cloned().unwrap_or_default();
            metadata.title.clone_from(&self.title_after);
            Some(metadata)
        } else {
            None
        };
        let meta_xml = requested_metadata
            .as_ref()
            .map(|metadata| crate::edit_ops::stage_metadata(self.source, metadata))
            .transpose()?;
        let staged_links = link_changes
            .iter()
            .map(|change| (change.reference, change.after.clone()))
            .collect::<Vec<_>>();
        let parts = crate::edit_ops::mutate_xml(
            self.source,
            &staged_links,
            &self.sections,
            &self.generated_indexes,
            &self.body_items,
            &self.styles,
        )?;
        let mut removed_resources = Vec::new();
        let mut resource_writes = Vec::new();
        removed_resources
            .try_reserve(self.resources.len())
            .map_err(|source| Error::Allocation {
                resource: "ODM removed resources",
                source,
            })?;
        resource_writes
            .try_reserve(self.resources.len())
            .map_err(|source| Error::Allocation {
                resource: "ODM resource writes",
                source,
            })?;
        for change in self.resources.values() {
            match change {
                ResourceChange::Put(resource) => {
                    resource_writes.push(crate::package::ResourceWrite {
                        path: resource.path().to_owned(),
                        media_type: resource.media_type().to_owned(),
                        bytes: resource.bytes().to_vec(),
                    });
                },
                ResourceChange::Remove(resource) => {
                    removed_resources.push(resource.path().to_owned());
                },
            }
        }
        ensure_removed_resources_are_unreferenced(&parts.content, &removed_resources)?;
        let snapshot = self.source.with_transaction_parts(
            &parts.content,
            parts.styles.as_deref(),
            meta_xml.as_deref(),
            &removed_resources,
            &resource_writes,
        )?;
        validate_security_policy(&snapshot, self.policy)?;
        if snapshot.title() != self.title_after.as_deref() {
            return Err(invalid(
                "ODM transaction title readback differs from the request",
            ));
        }
        for change in &link_changes {
            let actual = snapshot
                .subdocuments()
                .get(change.reference.get())
                .ok_or_else(|| invalid("ODM transaction link disappeared during readback"))?;
            if actual.href() != change.after {
                return Err(invalid(
                    "ODM transaction link readback differs from the request",
                ));
            }
        }
        if let Some(requested) = &requested_metadata {
            let actual = snapshot
                .metadata()
                .ok_or_else(|| invalid("ODM transaction metadata disappeared during readback"))?;
            if !simple_metadata_equal(actual, requested) {
                return Err(invalid(
                    "ODM transaction metadata readback differs from the request",
                ));
            }
        }
        verify_extended_readback(
            self.source,
            &snapshot,
            &self.sections,
            &self.generated_indexes,
            &self.body_items,
            &self.styles,
            &self.resources,
        )?;
        let title = title_changed.then_some(TitleChange {
            before: self.title_before,
            after: self.title_after,
        });
        Ok(Commit::new(
            self.source,
            snapshot,
            ChangeSet {
                title,
                links: link_changes,
                metadata: if explicit_metadata {
                    requested_metadata.map(|after| MetadataChange {
                        before: self.source.metadata().cloned().unwrap_or_default(),
                        after,
                    })
                } else {
                    None
                },
                sections: self.sections,
                generated_indexes: self.generated_indexes,
                body_items: self.body_items,
                styles: self.styles,
                resources: self.resources.into_values().collect(),
            },
        ))
    }
}

/// The title effect of a unified transaction.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TitleChange {
    before: Option<String>,
    after: Option<String>,
}

impl TitleChange {
    /// Returns the source title.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.before.as_deref()
    }

    /// Returns the published title.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }
}

/// One linked-section effect of a unified transaction.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LinkChange {
    reference: Position,
    before: String,
    after: String,
}

impl LinkChange {
    /// Returns the checked reference position.
    #[must_use]
    pub const fn reference(&self) -> Position {
        self.reference
    }

    /// Returns the source target.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// Returns the published target.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// Ordered semantic effects retained by a unified patch.
#[derive(Clone, Debug, Default)]
pub struct ChangeSet {
    title: Option<TitleChange>,
    links: Vec<LinkChange>,
    metadata: Option<MetadataChange>,
    sections: Vec<SectionChange>,
    generated_indexes: Vec<GeneratedIndexChange>,
    body_items: Vec<BodyItemChange>,
    styles: Vec<StyleChange>,
    resources: Vec<ResourceChange>,
}

impl ChangeSet {
    /// Returns the title effect, when present.
    #[must_use]
    pub const fn title(&self) -> Option<&TitleChange> {
        self.title.as_ref()
    }

    /// Returns link effects in semantic reference order.
    #[must_use]
    pub fn links(&self) -> &[LinkChange] {
        &self.links
    }

    /// Returns the simple metadata effect, when present.
    #[must_use]
    pub const fn metadata(&self) -> Option<&MetadataChange> {
        self.metadata.as_ref()
    }

    /// Returns section-tree effects in staging order.
    #[must_use]
    pub fn sections(&self) -> &[SectionChange] {
        &self.sections
    }

    /// Returns generated-index effects in staging order.
    #[must_use]
    pub fn generated_indexes(&self) -> &[GeneratedIndexChange] {
        &self.generated_indexes
    }

    /// Returns direct master-body item effects in staging order.
    #[must_use]
    pub fn body_items(&self) -> &[BodyItemChange] {
        &self.body_items
    }

    /// Returns style-catalog effects in staging order.
    #[must_use]
    pub fn styles(&self) -> &[StyleChange] {
        &self.styles
    }

    /// Returns resource effects in path order.
    #[must_use]
    pub fn resources(&self) -> &[ResourceChange] {
        &self.resources
    }

    /// Returns whether this set contains no semantic effect.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.title.is_none()
            && self.links.is_empty()
            && self.metadata.is_none()
            && self.sections.is_empty()
            && self.generated_indexes.is_empty()
            && self.body_items.is_empty()
            && self.styles.is_empty()
            && self.resources.is_empty()
    }
}

/// Before/after common metadata retained by a transaction.
#[derive(Clone, Debug)]
pub struct MetadataChange {
    before: litchi_core::Metadata,
    after: litchi_core::Metadata,
}

impl MetadataChange {
    /// Returns the source metadata projection.
    #[must_use]
    pub const fn before(&self) -> &litchi_core::Metadata {
        &self.before
    }

    /// Returns the published metadata projection.
    #[must_use]
    pub const fn after(&self) -> &litchi_core::Metadata {
        &self.after
    }
}

/// A fully validated unified transaction result.
pub struct Commit {
    snapshot: Master,
    patch: Patch,
}

impl Commit {
    fn new(source: &Master, snapshot: Master, changes: ChangeSet) -> Self {
        Self {
            patch: Patch {
                before: source.shared_bytes(),
                after: snapshot.shared_bytes(),
                changes,
            },
            snapshot,
        }
    }

    /// Returns the published immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Master {
        &self.snapshot
    }

    /// Returns the exact-source-checked reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the commit and returns its snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Master {
        self.snapshot
    }
}

/// An exact-source-checked reversible unified patch.
#[derive(Clone)]
pub struct Patch {
    before: Arc<Vec<u8>>,
    after: Arc<Vec<u8>>,
    changes: ChangeSet,
}

impl Patch {
    /// Returns the semantic effects.
    #[must_use]
    pub const fn changes(&self) -> &ChangeSet {
        &self.changes
    }

    /// Returns whether this patch applies to the exact source artifact.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Master) -> bool {
        source.as_bytes() == self.before.as_slice()
    }

    /// Applies this patch only to its exact immutable source.
    ///
    /// # Errors
    ///
    /// Returns an error when `source` differs byte-for-byte.
    pub fn apply(&self, source: &Master) -> Result<Master> {
        if !self.is_applicable_to(source) {
            return Err(invalid("ODM unified patch source does not match"));
        }
        Master::from_shared_bytes(Arc::clone(&self.after))
    }

    /// Returns a patch restoring the exact source package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            changes: inverse_changes(&self.changes),
        }
    }

    /// Returns whether the exact package bytes are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.as_slice() == self.after.as_slice()
    }

    /// Merges effects derived from the same source when their writes agree.
    ///
    /// # Errors
    ///
    /// Returns typed conflicts for overlapping divergent writes, a source
    /// mismatch, or a validation failure while publishing the merged package.
    pub fn merge(&self, other: &Self) -> std::result::Result<Self, MergeError> {
        if self.before.as_slice() != other.before.as_slice() {
            return Err(MergeError::DifferentSource);
        }
        let source =
            Master::from_shared_bytes(Arc::clone(&self.before)).map_err(MergeError::Invalid)?;
        let conflicts = find_conflicts(&self.changes, &other.changes, source.section_tree())
            .map_err(MergeError::Invalid)?;
        if !conflicts.is_empty() {
            return Err(MergeError::Conflicts(ConflictSet::new(conflicts)));
        }
        let merge_policy = if source.security().active_content().is_empty() {
            SecurityPolicy::default()
        } else {
            SecurityPolicy::default().with_active_content(ActiveContentPolicy::PreserveInert)
        };
        let mut edit = source.edit_with_policy(merge_policy);
        stage_changes(&mut edit, &self.changes).map_err(MergeError::Invalid)?;
        stage_changes(&mut edit, &other.changes).map_err(MergeError::Invalid)?;
        edit.commit()
            .map(|commit| commit.patch)
            .map_err(MergeError::Invalid)
    }

    /// Builds a non-mutating same-base three-way plan.
    ///
    /// The exact common source is the base and the two patches are branches.
    /// Planning never publishes package bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the branches do not share an exact base, the base
    /// cannot be reopened, or conflict discovery cannot allocate.
    pub fn plan_three_way(&self, other: &Self) -> std::result::Result<MergePlan, MergeError> {
        if self.before.as_slice() != other.before.as_slice() {
            return Err(MergeError::DifferentSource);
        }
        let source =
            Master::from_shared_bytes(Arc::clone(&self.before)).map_err(MergeError::Invalid)?;
        let conflicts = find_conflicts(&self.changes, &other.changes, source.section_tree())
            .map_err(MergeError::Invalid)?;
        Ok(MergePlan {
            left: self.clone(),
            right: other.clone(),
            conflicts: ConflictSet::new(conflicts),
        })
    }

    /// Converts this patch to bounded canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns an error if either exact package exceeds the durable bounds or
    /// cannot be reopened without credentials.
    pub fn durable(&self) -> Result<DurablePatch> {
        DurablePatch::from_artifacts(self.before.as_slice(), self.after.as_slice())
    }
}

/// Non-mutating three-way plan over one exact source and two branches.
pub struct MergePlan {
    left: Patch,
    right: Patch,
    conflicts: ConflictSet<Conflict>,
}

impl MergePlan {
    /// Returns deterministic overlap details without publishing a candidate.
    #[must_use]
    pub const fn conflicts(&self) -> &ConflictSet<Conflict> {
        &self.conflicts
    }

    /// Returns whether automatic commit is currently possible.
    #[must_use]
    pub fn can_commit(&self) -> bool {
        self.conflicts.is_empty()
    }

    /// Publishes the already-planned disjoint effects with full reopen.
    ///
    /// # Errors
    ///
    /// Returns the retained conflicts or a candidate validation failure.
    pub fn commit(self) -> std::result::Result<Patch, MergeError> {
        if !self.conflicts.is_empty() {
            return Err(MergeError::Conflicts(self.conflicts));
        }
        self.left.merge(&self.right)
    }
}

/// One semantic write target which prevented an automatic merge.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Conflict {
    /// Both patches write different titles.
    Title,
    /// Both patches write one link to different targets.
    Link(Position),
    /// Both patches write simple metadata differently.
    Metadata,
    /// Both patches divergently write one supported metadata field.
    MetadataField(&'static str),
    /// Both patches structurally write one source section.
    Section(Position),
    /// Both patches create divergent effects at one section destination name.
    SectionName(String),
    /// Both patches divergently write one generated index.
    GeneratedIndex(Position),
    /// Both patches rename generated indexes to the same identity.
    GeneratedIndexName(String),
    /// Both patches divergently write one direct master-body item.
    BodyItem(Position),
    /// Both patches create a generated index with the same identity.
    BodyItemName(String),
    /// Both patches write one style identity.
    Style(String),
    /// Both patches write one package resource path.
    Resource(String),
}

/// A unified-patch merge failure.
#[derive(Debug)]
#[non_exhaustive]
pub enum MergeError {
    /// The patches do not share exact source bytes.
    DifferentSource,
    /// Divergent writes overlap.
    Conflicts(ConflictSet<Conflict>),
    /// The merged candidate failed validation.
    Invalid(Error),
}

/// Explicit bounded undo/redo history for master snapshots.
pub struct History {
    inner: CoreHistory<Master>,
}

impl History {
    pub(crate) fn new(current: Master, limits: HistoryLimits) -> Self {
        Self {
            inner: CoreHistory::new(current, limits),
        }
    }

    /// Returns the current immutable snapshot.
    #[must_use]
    pub const fn current(&self) -> &Master {
        self.inner.current()
    }

    /// Records a commit only when history currently points at its source.
    ///
    /// # Errors
    ///
    /// Returns an error for stale lineage, size overflow, or exceeded bounds.
    pub fn record(&mut self, commit: &Commit) -> Result<Vec<Master>> {
        if !commit.patch.is_applicable_to(self.current()) {
            return Err(invalid(
                "ODM history commit source does not match current state",
            ));
        }
        let weight = u64::try_from(commit.patch.before.len())
            .ok()
            .and_then(|before| {
                u64::try_from(commit.patch.after.len())
                    .ok()
                    .and_then(|after| before.checked_add(after))
            })
            .ok_or_else(|| invalid("ODM history transition weight overflow"))?;
        self.inner
            .record(commit.snapshot.clone(), weight)
            .map_err(patch_wire_error)
    }

    /// Moves to the previous retained snapshot.
    pub fn undo(&mut self) -> bool {
        self.inner.undo()
    }

    /// Moves to the next retained snapshot.
    pub fn redo(&mut self) -> bool {
        self.inner.redo()
    }

    /// Returns whether undo is available.
    #[must_use]
    pub fn can_undo(&self) -> bool {
        self.inner.can_undo()
    }

    /// Returns whether redo is available.
    #[must_use]
    pub fn can_redo(&self) -> bool {
        self.inner.can_redo()
    }
}

/// Bounded reversible durable ODM patch.
#[derive(Clone)]
pub struct DurablePatch {
    inner: CorePatch<Reversible>,
}

impl DurablePatch {
    fn from_artifacts(before: &[u8], after: &[u8]) -> Result<Self> {
        Master::from_bytes(copy_bytes(before)?)?;
        Master::from_bytes(copy_bytes(after)?)?;
        let limits = durable_limits();
        let mut forward_blobs = BlobBundle::new(limits.blobs());
        let after_id = forward_blobs.insert(after).map_err(patch_wire_error)?;
        let mut reverse_blobs = BlobBundle::new(limits.blobs());
        let before_id = reverse_blobs.insert(before).map_err(patch_wire_error)?;
        let forward = durable_operation(limits, &before_id, &after_id)?;
        let inverse = durable_operation(limits, &after_id, &before_id)?;
        let inner = CorePatch::<Reversible>::new(
            limits,
            DURABLE_FORMAT,
            [ReversibleOperation::new(forward, inverse)],
            forward_blobs,
            reverse_blobs,
        )
        .map_err(patch_wire_error)?;
        Ok(Self { inner })
    }

    /// Parses canonical deterministic JSON and validates both ODM artifacts.
    ///
    /// # Errors
    ///
    /// Returns an error for non-canonical, foreign, excessive, or invalid data.
    pub fn from_deterministic_json(bytes: &[u8]) -> Result<Self> {
        let inner = CorePatch::<Reversible>::from_deterministic_json(bytes, durable_limits())
            .map_err(patch_wire_error)?;
        validate_reversible(&inner)?;
        Ok(Self { inner })
    }

    /// Serializes canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns an error when bounded serialization fails.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>> {
        self.inner.to_deterministic_json().map_err(patch_wire_error)
    }

    /// Applies this durable patch to its exact source artifact.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source bytes or invalid target bytes.
    pub fn apply(&self, source: &Master) -> Result<Master> {
        let inverse = self.inner.inverse();
        if source.as_bytes() != durable_direction(&inverse)?.target_bytes {
            return Err(invalid("ODM durable patch source does not match"));
        }
        master_from_target(&self.inner)
    }

    /// Returns the exact durable inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            inner: self.inner.inverse(),
        }
    }

    /// Permanently removes inverse material.
    #[must_use]
    pub fn seal(self) -> SealedPatch {
        SealedPatch {
            inner: self.inner.seal(),
        }
    }
}

/// Forward-only durable ODM patch.
#[derive(Clone)]
pub struct SealedPatch {
    inner: CorePatch<ForwardOnly>,
}

impl SealedPatch {
    /// Parses a canonical forward-only durable patch.
    ///
    /// # Errors
    ///
    /// Returns an error for non-canonical, foreign, excessive, or invalid data.
    pub fn from_deterministic_json(bytes: &[u8]) -> Result<Self> {
        let inner = CorePatch::<ForwardOnly>::from_deterministic_json(bytes, durable_limits())
            .map_err(patch_wire_error)?;
        validate_sealed(&inner)?;
        Ok(Self { inner })
    }

    /// Serializes canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns an error when bounded serialization fails.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>> {
        self.inner.to_deterministic_json().map_err(patch_wire_error)
    }

    /// Applies after checking the retained SHA-256 source precondition.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source bytes or invalid target bytes.
    pub fn apply(&self, source: &Master) -> Result<Master> {
        let direction = durable_direction(&self.inner)?;
        if BlobId::of(source.as_bytes()).as_hex() != direction.source_id {
            return Err(invalid("ODM durable patch source does not match"));
        }
        Master::from_bytes(copy_bytes(direction.target_bytes)?)
    }
}

fn resolve(source: &Master, selector: Selector<'_>) -> Result<Position> {
    match selector {
        Selector::Position(position) => source
            .subdocuments()
            .get(position.get())
            .map(|_| position)
            .ok_or_else(|| invalid("ODM subdocument selector is out of bounds")),
        Selector::Section(name) => source
            .subdocuments()
            .iter()
            .position(|reference| reference.section() == name.as_ref())
            .map(Position::new)
            .ok_or_else(|| invalid("ODM linked section name was not found")),
    }
}

fn is_external_target(href: &str) -> bool {
    href.is_empty()
        || href.starts_with('/')
        || href.starts_with('\\')
        || href.contains('\\')
        || href.contains(':')
        || href
            .split('/')
            .any(|segment| segment.is_empty() || matches!(segment, "." | ".."))
}

fn section_name_exists(source: &Master, staged: &[SectionChange], name: &str) -> bool {
    source
        .section_tree()
        .sections()
        .iter()
        .any(|section| section.name() == name)
        || staged.iter().any(|change| match change {
            SectionChange::Add(section) => section.name() == name,
            SectionChange::Rename { after, .. } => after == name,
            SectionChange::Remove { .. } => false,
        })
}

fn generated_index_name_exists(
    source: &Master,
    generated_indexes: &[GeneratedIndexChange],
    body_items: &[BodyItemChange],
    name: &str,
) -> bool {
    source
        .structure()
        .generated_indexes()
        .iter()
        .any(|index| index.name() == Some(name))
        || generated_indexes.iter().any(|change| {
            matches!(change, GeneratedIndexChange::Rename { after, .. } if after == name)
        })
        || body_items.iter().any(|change| {
            matches!(change, BodyItemChange::Add(spec) if spec.generated_index_name() == Some(name))
        })
}

fn ensure_section_not_staged(staged: &[SectionChange], position: Position) -> Result<()> {
    if staged.iter().any(|change| match change {
        SectionChange::Rename {
            position: selected, ..
        }
        | SectionChange::Remove {
            position: selected, ..
        } => *selected == position,
        SectionChange::Add(_) => false,
    }) {
        return Err(invalid(
            "ODM section already has a staged structural intent",
        ));
    }
    Ok(())
}

fn style_name_exists(source: &Master, staged: &[StyleChange], name: &str) -> bool {
    source.styles().iter().any(|style| style.name() == name)
        || staged.iter().any(|change| match change {
            StyleChange::Add(style) => style.name() == name,
            StyleChange::Rename { after, .. } => after == name,
            StyleChange::Remove { .. } => false,
        })
}

fn transfer_style_closure(
    edit: &mut Edit<'_>,
    source: &Master,
    leaf_name: &str,
    collision: CollisionPolicy,
) -> Result<String> {
    let mut closure = Vec::new();
    let mut seen = HashSet::new();
    let mut current = Some(leaf_name.to_string());
    while let Some(name) = current {
        if !seen.insert(name.clone()) {
            return Err(invalid("ODM transfer style parent cycle is not supported"));
        }
        let matches = source
            .styles()
            .iter()
            .filter(|definition| definition.name() == name)
            .collect::<Vec<_>>();
        let definition = match matches.as_slice() {
            [definition] => *definition,
            [] => return Err(invalid("ODM transfer source style dependency is missing")),
            _ => return Err(invalid("ODM transfer source style dependency is ambiguous")),
        };
        current = definition.parent().map(str::to_string);
        closure.push(definition);
    }
    closure.reverse();

    let mut names = HashMap::new();
    for definition in closure {
        let mapped_parent = definition
            .parent()
            .map(|parent| {
                names
                    .get(parent)
                    .cloned()
                    .ok_or_else(|| invalid("ODM transferred style parent mapping is incomplete"))
            })
            .transpose()?;
        let desired = definition.name();
        let source_xml = style_owner_xml(source, definition)?;
        let desired_spec =
            StyleSpec::imported(source_xml, definition, desired, mapped_parent.as_deref())?;
        let occupied = style_name_exists(edit.source, &edit.styles, desired);
        let identical = occupied
            && transferred_style_is_identical(
                edit.source,
                &edit.styles,
                source,
                definition,
                &desired_spec,
            )?;
        let destination_name = if !occupied {
            desired.to_string()
        } else if identical && collision != CollisionPolicy::Refuse {
            names.insert(desired.to_string(), desired.to_string());
            continue;
        } else if collision == CollisionPolicy::Rename {
            unique_style_name(edit.source, &edit.styles, desired)?
        } else {
            return Err(invalid(
                "ODM linked-section transfer style identity collision",
            ));
        };
        let spec = if destination_name == desired {
            desired_spec
        } else {
            StyleSpec::imported(
                source_xml,
                definition,
                &destination_name,
                mapped_parent.as_deref(),
            )?
        };
        edit.add_style(spec)?;
        names.insert(desired.to_string(), destination_name);
    }
    names
        .remove(leaf_name)
        .ok_or_else(|| invalid("ODM transferred leaf style mapping disappeared"))
}

fn style_owner_xml<'source>(
    source: &'source Master,
    definition: &crate::style::Definition,
) -> Result<&'source str> {
    match definition.origin() {
        crate::style::Origin::Content => Ok(source.content_xml()),
        crate::style::Origin::Styles => source
            .styles_xml()
            .ok_or_else(|| invalid("ODM transferred style owner is missing")),
    }
}

fn transferred_style_is_identical(
    destination: &Master,
    staged: &[StyleChange],
    source: &Master,
    source_definition: &crate::style::Definition,
    desired_spec: &StyleSpec,
) -> Result<bool> {
    if staged.iter().any(|change| {
        matches!(change, StyleChange::Add(spec)
            if spec.name() == desired_spec.name()
                && spec.family() == desired_spec.family()
                && spec.origin() == desired_spec.origin()
                && spec.parent() == desired_spec.parent()
                && spec.raw_fragment() == desired_spec.raw_fragment())
    }) {
        return Ok(true);
    }
    let matches = destination
        .styles()
        .iter()
        .filter(|definition| definition.name() == source_definition.name())
        .collect::<Vec<_>>();
    let [destination_definition] = matches.as_slice() else {
        return Ok(false);
    };
    if destination_definition.origin() != source_definition.origin()
        || destination_definition.family() != source_definition.family()
        || destination_definition.parent() != source_definition.parent()
    {
        return Ok(false);
    }
    let source_fragment = style_owner_xml(source, source_definition)?
        .get(source_definition.source_span.clone())
        .ok_or_else(|| invalid("ODM source style span is stale"))?;
    let destination_fragment = style_owner_xml(destination, destination_definition)?
        .get(destination_definition.source_span.clone())
        .ok_or_else(|| invalid("ODM destination style span is stale"))?;
    Ok(source_fragment == destination_fragment)
}

fn unique_style_name(source: &Master, staged: &[StyleChange], base: &str) -> Result<String> {
    for sequence in 1..=1_000_000usize {
        let candidate = format!("{base}__import{sequence}");
        crate::edit_ops::validate_value(&candidate, "ODM imported style name", false)?;
        if !style_name_exists(source, staged, &candidate) {
            return Ok(candidate);
        }
    }
    Err(invalid(
        "ODM imported style collision sequence is exhausted",
    ))
}

fn resolve_resource_collision(
    destination: &Master,
    staged: &BTreeMap<String, ResourceChange>,
    desired: &str,
    media_type: &str,
    bytes: &[u8],
    collision: CollisionPolicy,
) -> Result<(String, bool)> {
    ResourceSpec::new(desired, media_type, Vec::new())?;
    let mut candidate = desired.to_string();
    for sequence in 0..=1_000_000usize {
        let existing = staged_resource(staged, &candidate).or_else(|| {
            destination
                .resources()
                .resources()
                .iter()
                .find(|resource| resource.path() == candidate)
                .map(|resource| (resource.media_type(), None))
        });
        let Some((existing_media_type, staged_bytes)) = existing else {
            return Ok((candidate, true));
        };
        let existing_bytes = staged_bytes.map_or_else(
            || destination.resource_bytes(&candidate),
            |resource_bytes| Ok(resource_bytes.to_owned()),
        )?;
        let identical = existing_media_type.unwrap_or("application/octet-stream") == media_type
            && existing_bytes == bytes;
        if identical && collision != CollisionPolicy::Refuse {
            return Ok((candidate, false));
        }
        if collision != CollisionPolicy::Rename {
            return Err(invalid("ODM transfer resource destination already exists"));
        }
        candidate = imported_resource_path(desired, sequence.saturating_add(1))?;
    }
    Err(invalid(
        "ODM imported resource collision sequence is exhausted",
    ))
}

fn staged_resource<'change>(
    staged: &'change BTreeMap<String, ResourceChange>,
    path: &str,
) -> Option<(Option<&'change str>, Option<&'change [u8]>)> {
    match staged.get(path) {
        Some(ResourceChange::Put(resource)) => {
            Some((Some(resource.media_type()), Some(resource.bytes())))
        },
        Some(ResourceChange::Remove(_)) | None => None,
    }
}

fn imported_resource_path(path: &str, sequence: usize) -> Result<String> {
    let (directory, file_name) = path
        .rsplit_once('/')
        .map_or(("", path), |(dir, name)| (dir, name));
    let (stem, extension) = file_name
        .rsplit_once('.')
        .map_or((file_name, ""), |(stem, extension)| (stem, extension));
    let imported_file_name = if extension.is_empty() {
        format!("{stem}__import{sequence}")
    } else {
        format!("{stem}__import{sequence}.{extension}")
    };
    let candidate = if directory.is_empty() {
        imported_file_name
    } else {
        format!("{directory}/{imported_file_name}")
    };
    ResourceSpec::new(&candidate, "application/octet-stream", Vec::new())?;
    Ok(candidate)
}

fn resolve_style(source: &Master, origin: crate::style::Origin, name: &str) -> Result<()> {
    let count = source
        .styles()
        .iter()
        .filter(|style| style.origin() == origin && style.name() == name)
        .count();
    match count {
        1 => Ok(()),
        0 => Err(invalid("ODM style selector was not found")),
        _ => Err(invalid("ODM style selector is ambiguous")),
    }
}

fn ensure_removed_resources_are_unreferenced(content: &str, removed: &[String]) -> Result<()> {
    if removed.is_empty() {
        return Ok(());
    }
    let semantics = crate::codec::parse(content)?;
    if semantics
        .references()
        .iter()
        .any(|reference| removed.iter().any(|path| path == reference.href()))
    {
        return Err(invalid(
            "ODM resource removal is blocked by a linked-section dependency",
        ));
    }
    Ok(())
}

fn validate_security_policy(snapshot: &Master, policy: SecurityPolicy) -> Result<()> {
    if policy.active_content() == ActiveContentPolicy::Refuse
        && !snapshot.security().active_content().is_empty()
    {
        return Err(invalid(
            "ODM security policy refuses changed output containing active content",
        ));
    }
    if !policy.allows_external_targets()
        && snapshot
            .subdocuments()
            .iter()
            .any(|reference| reference.target().is_external())
    {
        return Err(invalid(
            "ODM security policy refuses final external link targets",
        ));
    }
    if !policy.allows_missing_package_targets() && !snapshot.resources().missing().is_empty() {
        return Err(invalid(
            "ODM security policy refuses unresolved package targets",
        ));
    }
    Ok(())
}

fn simple_metadata_equal(left: &litchi_core::Metadata, right: &litchi_core::Metadata) -> bool {
    left.title == right.title
        && left.author == right.author
        && left.subject == right.subject
        && left.description == right.description
        && left.keywords == right.keywords
}

fn verify_extended_readback(
    source: &Master,
    snapshot: &Master,
    sections: &[SectionChange],
    generated_indexes: &[GeneratedIndexChange],
    body_items: &[BodyItemChange],
    styles: &[StyleChange],
    resources: &BTreeMap<String, ResourceChange>,
) -> Result<()> {
    for change in sections {
        match change {
            SectionChange::Add(section) => {
                let Some(node) = snapshot
                    .section_tree()
                    .sections()
                    .iter()
                    .find(|node| node.name() == section.name())
                else {
                    return Err(invalid("ODM added section failed semantic readback"));
                };
                if let Some(expected) = section.subdocument() {
                    let actual = node
                        .reference()
                        .and_then(|position| snapshot.subdocuments().get(position.get()))
                        .ok_or_else(|| invalid("ODM added linked section failed readback"))?;
                    if actual.href() != expected.href()
                        || actual.source_section() != expected.source_section()
                        || actual.filter_name() != expected.filter_name()
                    {
                        return Err(invalid(
                            "ODM added linked-section semantics differ from the request",
                        ));
                    }
                }
            },
            SectionChange::Rename { before, after, .. } => {
                let names = snapshot.section_tree().sections();
                if names.iter().any(|node| node.name() == before)
                    || !names.iter().any(|node| node.name() == after)
                {
                    return Err(invalid("ODM renamed section failed semantic readback"));
                }
            },
            SectionChange::Remove { before, .. } => {
                if snapshot
                    .section_tree()
                    .sections()
                    .iter()
                    .any(|node| node.name() == before)
                {
                    return Err(invalid("ODM removed section failed semantic readback"));
                }
            },
        }
    }
    for change in generated_indexes {
        match change {
            GeneratedIndexChange::Rename { after, .. } => {
                if !snapshot
                    .structure()
                    .generated_indexes()
                    .iter()
                    .any(|index| index.name() == Some(after))
                {
                    return Err(invalid("ODM generated-index rename readback differs"));
                }
            },
        }
    }
    for change in body_items {
        let kind = match change {
            BodyItemChange::Add(spec) => spec.kind(),
            BodyItemChange::Remove { kind, .. } => *kind,
        };
        let added = body_items
            .iter()
            .filter(
                |candidate| matches!(candidate, BodyItemChange::Add(spec) if spec.kind() == kind),
            )
            .count();
        let removed = body_items
            .iter()
            .filter(|candidate| {
                matches!(candidate, BodyItemChange::Remove { kind: candidate_kind, .. } if *candidate_kind == kind)
            })
            .count();
        let before = source
            .structure()
            .items()
            .iter()
            .filter(|candidate| **candidate == kind)
            .count();
        let expected = before
            .checked_add(added)
            .ok_or_else(|| invalid("ODM body-item addition inventory overflowed"))?
            .checked_sub(removed)
            .ok_or_else(|| invalid("ODM body-item inventory is inconsistent"))?;
        let actual = snapshot
            .structure()
            .items()
            .iter()
            .filter(|candidate| **candidate == kind)
            .count();
        if actual != expected {
            return Err(invalid("ODM body-item change failed semantic readback"));
        }
        if let BodyItemChange::Add(spec) = change
            && let Some(name) = spec.generated_index_name()
            && !snapshot
                .structure()
                .generated_indexes()
                .iter()
                .any(|index| index.name() == Some(name))
        {
            return Err(invalid(
                "ODM added generated index failed semantic readback",
            ));
        }
    }
    for change in styles {
        match change {
            StyleChange::Add(style) => {
                let definition = snapshot
                    .styles()
                    .iter()
                    .find(|definition| {
                        definition.origin() == style.origin() && definition.name() == style.name()
                    })
                    .ok_or_else(|| invalid("ODM added style failed semantic readback"))?;
                if definition.family() != Some(style.family())
                    || definition.parent() != style.parent()
                {
                    return Err(invalid("ODM added style dependency readback differs"));
                }
                if let Some(expected) = style.raw_fragment() {
                    let actual = style_owner_xml(snapshot, definition)?
                        .get(definition.source_span.clone())
                        .ok_or_else(|| invalid("ODM added style readback span is stale"))?;
                    if actual != expected {
                        return Err(invalid("ODM imported style XML readback differs"));
                    }
                }
            },
            StyleChange::Rename {
                origin,
                before,
                after,
            } => {
                if snapshot
                    .styles()
                    .iter()
                    .any(|definition| definition.origin() == *origin && definition.name() == before)
                    || !snapshot.styles().iter().any(|definition| {
                        definition.origin() == *origin && definition.name() == after
                    })
                {
                    return Err(invalid("ODM renamed style failed semantic readback"));
                }
            },
            StyleChange::Remove { origin, name } => {
                if snapshot
                    .styles()
                    .iter()
                    .any(|definition| definition.origin() == *origin && definition.name() == name)
                {
                    return Err(invalid("ODM removed style failed semantic readback"));
                }
            },
        }
    }
    for (path, change) in resources {
        let present = snapshot
            .resources()
            .resources()
            .iter()
            .any(|resource| resource.path() == path);
        if matches!(change, ResourceChange::Put(_)) != present {
            return Err(invalid("ODM resource operation failed semantic readback"));
        }
    }
    Ok(())
}

fn collect_link_changes(
    source: &Master,
    staged: &BTreeMap<usize, String>,
) -> Result<Vec<LinkChange>> {
    let mut changes = Vec::new();
    changes
        .try_reserve(staged.len())
        .map_err(|allocation| Error::Allocation {
            resource: "ODM unified link changes",
            source: allocation,
        })?;
    for (&index, after) in staged {
        let before = source
            .subdocuments()
            .get(index)
            .ok_or_else(|| invalid("ODM staged reference disappeared"))?
            .href();
        if before != after {
            changes.push(LinkChange {
                reference: Position::new(index),
                before: before.to_owned(),
                after: after.clone(),
            });
        }
    }
    Ok(changes)
}

fn inverse_changes(changes: &ChangeSet) -> ChangeSet {
    ChangeSet {
        title: changes.title.as_ref().map(|change| TitleChange {
            before: change.after.clone(),
            after: change.before.clone(),
        }),
        links: changes
            .links
            .iter()
            .map(|change| LinkChange {
                reference: change.reference,
                before: change.after.clone(),
                after: change.before.clone(),
            })
            .collect(),
        metadata: changes.metadata.as_ref().map(|change| MetadataChange {
            before: change.after.clone(),
            after: change.before.clone(),
        }),
        sections: changes
            .sections
            .iter()
            .filter_map(|change| match change {
                SectionChange::Rename {
                    position,
                    before,
                    after,
                } => Some(SectionChange::Rename {
                    position: *position,
                    before: after.clone(),
                    after: before.clone(),
                }),
                SectionChange::Add(_) | SectionChange::Remove { .. } => None,
            })
            .collect(),
        generated_indexes: changes
            .generated_indexes
            .iter()
            .map(|change| match change {
                GeneratedIndexChange::Rename {
                    item,
                    before,
                    after,
                } => GeneratedIndexChange::Rename {
                    item: *item,
                    before: after.clone(),
                    after: before.clone(),
                },
            })
            .collect(),
        body_items: Vec::new(),
        styles: changes
            .styles
            .iter()
            .filter_map(|change| match change {
                StyleChange::Rename {
                    origin,
                    before,
                    after,
                } => Some(StyleChange::Rename {
                    origin: *origin,
                    before: after.clone(),
                    after: before.clone(),
                }),
                StyleChange::Add(_) | StyleChange::Remove { .. } => None,
            })
            .collect(),
        resources: changes
            .resources
            .iter()
            .map(|change| match change {
                ResourceChange::Put(resource) => ResourceChange::Remove(resource.clone()),
                ResourceChange::Remove(resource) => ResourceChange::Put(resource.clone()),
            })
            .collect(),
    }
}

fn find_conflicts(
    left: &ChangeSet,
    right: &ChangeSet,
    sections: &crate::section::Tree,
) -> Result<Vec<Conflict>> {
    let mut conflicts = Vec::new();
    let capacity = left
        .links
        .len()
        .min(right.links.len())
        .saturating_add(6)
        .saturating_add(left.sections.len().saturating_mul(right.sections.len()))
        .saturating_add(
            left.generated_indexes
                .len()
                .saturating_mul(right.generated_indexes.len()),
        )
        .saturating_add(left.body_items.len().saturating_mul(right.body_items.len()))
        .saturating_add(
            left.body_items
                .len()
                .saturating_mul(right.generated_indexes.len()),
        )
        .saturating_add(
            right
                .body_items
                .len()
                .saturating_mul(left.generated_indexes.len()),
        )
        .saturating_add(left.styles.len().saturating_mul(right.styles.len()))
        .saturating_add(left.resources.len().saturating_mul(right.resources.len()));
    conflicts
        .try_reserve(capacity)
        .map_err(|allocation| Error::Allocation {
            resource: "ODM merge conflicts",
            source: allocation,
        })?;
    if let (Some(left), Some(right)) = (&left.title, &right.title)
        && left.after != right.after
    {
        conflicts.push(Conflict::Title);
    }
    let mut right_links = HashMap::new();
    right_links
        .try_reserve(right.links.len())
        .map_err(|allocation| Error::Allocation {
            resource: "ODM merge link index",
            source: allocation,
        })?;
    for change in &right.links {
        right_links.insert(change.reference.get(), change.after.as_str());
    }
    for change in &left.links {
        if right_links
            .get(&change.reference.get())
            .is_some_and(|after| *after != change.after)
        {
            conflicts.push(Conflict::Link(change.reference));
        }
    }
    if let (Some(left), Some(right)) = (&left.metadata, &right.metadata) {
        collect_metadata_conflicts(left, right, &mut conflicts);
    }
    for left_change in &left.sections {
        for right_change in &right.sections {
            if let Some(position) = section_change_overlap(sections, left_change, right_change)
                && left_change != right_change
            {
                conflicts.push(Conflict::Section(position));
            }
            if let (Some(left_name), Some(right_name)) = (
                section_change_destination(left_change),
                section_change_destination(right_change),
            ) && left_name == right_name
                && left_change != right_change
            {
                conflicts.push(Conflict::SectionName(left_name.to_owned()));
            }
        }
    }
    for left_change in &left.generated_indexes {
        for right_change in &right.generated_indexes {
            let (
                GeneratedIndexChange::Rename {
                    item: left_item,
                    after: left_after,
                    ..
                },
                GeneratedIndexChange::Rename {
                    item: right_item,
                    after: right_after,
                    ..
                },
            ) = (left_change, right_change);
            if left_item == right_item && left_change != right_change {
                conflicts.push(Conflict::GeneratedIndex(*left_item));
            } else if left_item != right_item && left_after == right_after {
                conflicts.push(Conflict::GeneratedIndexName(left_after.clone()));
            }
        }
    }
    for left_change in &left.body_items {
        for right_change in &right.body_items {
            match (left_change, right_change) {
                (
                    BodyItemChange::Remove {
                        item: left_item, ..
                    },
                    BodyItemChange::Remove {
                        item: right_item, ..
                    },
                ) if left_item == right_item && left_change != right_change => {
                    conflicts.push(Conflict::BodyItem(*left_item));
                },
                (BodyItemChange::Add(left_spec), BodyItemChange::Add(right_spec))
                    if left_spec != right_spec
                        && left_spec.generated_index_name().is_some()
                        && left_spec.generated_index_name()
                            == right_spec.generated_index_name() =>
                {
                    conflicts.push(Conflict::BodyItemName(
                        left_spec
                            .generated_index_name()
                            .unwrap_or_default()
                            .to_owned(),
                    ));
                },
                _ => {},
            }
        }
        match left_change {
            BodyItemChange::Remove {
                item: left_item, ..
            } if right.generated_indexes.iter().any(|change| {
                matches!(change, GeneratedIndexChange::Rename { item, .. } if item == left_item)
            }) => conflicts.push(Conflict::BodyItem(*left_item)),
            BodyItemChange::Add(spec) => {
                if let Some(name) = spec.generated_index_name()
                    && right.generated_indexes.iter().any(|change| {
                        matches!(change, GeneratedIndexChange::Rename { after, .. } if after == name)
                    })
                {
                    conflicts.push(Conflict::BodyItemName(name.to_owned()));
                }
            },
            BodyItemChange::Remove { .. } => {},
        }
    }
    for right_change in &right.body_items {
        match right_change {
            BodyItemChange::Remove {
                item: right_item, ..
            } if left.generated_indexes.iter().any(|change| {
                matches!(change, GeneratedIndexChange::Rename { item, .. } if item == right_item)
            }) => conflicts.push(Conflict::BodyItem(*right_item)),
            BodyItemChange::Add(spec) => {
                if let Some(name) = spec.generated_index_name()
                    && left.generated_indexes.iter().any(|change| {
                        matches!(change, GeneratedIndexChange::Rename { after, .. } if after == name)
                    })
                {
                    conflicts.push(Conflict::BodyItemName(name.to_owned()));
                }
            },
            BodyItemChange::Remove { .. } => {},
        }
    }
    for left_change in &left.styles {
        for right_change in &right.styles {
            if style_change_key(left_change) == style_change_key(right_change)
                && left_change != right_change
            {
                conflicts.push(Conflict::Style(style_change_key(left_change).to_owned()));
            }
        }
    }
    for left_change in &left.resources {
        for right_change in &right.resources {
            if resource_change_path(left_change) == resource_change_path(right_change)
                && left_change != right_change
            {
                conflicts.push(Conflict::Resource(
                    resource_change_path(left_change).to_owned(),
                ));
            }
        }
    }
    Ok(conflicts)
}

fn collect_metadata_conflicts(
    left: &MetadataChange,
    right: &MetadataChange,
    conflicts: &mut Vec<Conflict>,
) {
    for (field, left_before, left_after, right_before, right_after) in [
        (
            "title",
            left.before.title.as_deref(),
            left.after.title.as_deref(),
            right.before.title.as_deref(),
            right.after.title.as_deref(),
        ),
        (
            "author",
            left.before.author.as_deref(),
            left.after.author.as_deref(),
            right.before.author.as_deref(),
            right.after.author.as_deref(),
        ),
        (
            "subject",
            left.before.subject.as_deref(),
            left.after.subject.as_deref(),
            right.before.subject.as_deref(),
            right.after.subject.as_deref(),
        ),
        (
            "description",
            left.before.description.as_deref(),
            left.after.description.as_deref(),
            right.before.description.as_deref(),
            right.after.description.as_deref(),
        ),
        (
            "keywords",
            left.before.keywords.as_deref(),
            left.after.keywords.as_deref(),
            right.before.keywords.as_deref(),
            right.after.keywords.as_deref(),
        ),
    ] {
        if left_before != left_after && right_before != right_after && left_after != right_after {
            conflicts.push(Conflict::MetadataField(field));
        }
    }
}

fn section_change_position(change: &SectionChange) -> Option<Position> {
    match change {
        SectionChange::Rename { position, .. } | SectionChange::Remove { position, .. } => {
            Some(*position)
        },
        SectionChange::Add(_) => None,
    }
}

fn section_change_overlap(
    tree: &crate::section::Tree,
    left: &SectionChange,
    right: &SectionChange,
) -> Option<Position> {
    let left_position = section_change_position(left)?;
    let right_position = section_change_position(right)?;
    if left_position == right_position {
        return Some(left_position);
    }
    let left_node = tree.get(left_position)?;
    let right_node = tree.get(right_position)?;
    if matches!(left, SectionChange::Remove { .. })
        && left_node.source_span.start <= right_node.source_span.start
        && right_node.source_span.end <= left_node.source_span.end
    {
        return Some(right_position);
    }
    if matches!(right, SectionChange::Remove { .. })
        && right_node.source_span.start <= left_node.source_span.start
        && left_node.source_span.end <= right_node.source_span.end
    {
        return Some(left_position);
    }
    None
}

fn section_change_destination(change: &SectionChange) -> Option<&str> {
    match change {
        SectionChange::Add(spec) => Some(spec.name()),
        SectionChange::Rename { after, .. } => Some(after),
        SectionChange::Remove { .. } => None,
    }
}

fn style_change_key(change: &StyleChange) -> &str {
    match change {
        StyleChange::Add(spec) => spec.name(),
        StyleChange::Rename { before, .. } => before,
        StyleChange::Remove { name, .. } => name,
    }
}

fn resource_change_path(change: &ResourceChange) -> &str {
    match change {
        ResourceChange::Put(spec) | ResourceChange::Remove(spec) => spec.path(),
    }
}

fn stage_changes(edit: &mut Edit<'_>, changes: &ChangeSet) -> Result<()> {
    if let Some(metadata) = &changes.metadata {
        stage_metadata_delta(edit, metadata)?;
    }
    if let Some(title) = &changes.title {
        if let Some(after) = &title.after {
            edit.set_title(after.clone())?;
        } else {
            edit.clear_title();
        }
    }
    for link in &changes.links {
        edit.set_link(link.reference, link.after.clone())?;
    }
    for generated_index in &changes.generated_indexes {
        if edit.generated_indexes.contains(generated_index) {
            continue;
        }
        match generated_index {
            GeneratedIndexChange::Rename { item, after, .. } => {
                edit.rename_generated_index(*item, after.clone())?;
            },
        }
    }
    for body_item in &changes.body_items {
        if edit.body_items.contains(body_item) {
            continue;
        }
        match body_item {
            BodyItemChange::Add(spec) => {
                edit.add_body_item(spec.clone())?;
            },
            BodyItemChange::Remove { item, .. } => {
                edit.remove_body_item(*item)?;
            },
        }
    }
    // Style additions/renames must be visible before strict section staging:
    // an added section may reference a style imported by the same patch.
    // Removals remain atomic because XML mutation receives both complete
    // intent lists and computes removed content spans before staging styles.
    for style in &changes.styles {
        if edit.styles.contains(style) {
            continue;
        }
        match style {
            StyleChange::Add(spec) => {
                edit.add_style(spec.clone())?;
            },
            StyleChange::Rename {
                origin,
                before,
                after,
            } => {
                edit.rename_style(*origin, before.clone(), after.clone())?;
            },
            StyleChange::Remove { origin, name } => {
                edit.remove_style(*origin, name.clone())?;
            },
        }
    }
    for section in &changes.sections {
        if edit.sections.contains(section) {
            continue;
        }
        match section {
            SectionChange::Add(spec) => {
                edit.add_section(spec.clone())?;
            },
            SectionChange::Rename {
                position, after, ..
            } => {
                edit.rename_section(*position, after.clone())?;
            },
            SectionChange::Remove { position, .. } => {
                edit.remove_section(*position)?;
            },
        }
    }
    for resource in &changes.resources {
        match resource {
            ResourceChange::Put(spec) => {
                edit.put_resource(spec.clone())?;
            },
            ResourceChange::Remove(spec) => {
                edit.remove_resource(spec.path().to_owned())?;
            },
        }
    }
    Ok(())
}

fn stage_metadata_delta(edit: &mut Edit<'_>, change: &MetadataChange) -> Result<()> {
    let mut target = edit
        .metadata
        .clone()
        .unwrap_or_else(|| edit.source.metadata().cloned().unwrap_or_default());
    if change.before.title == change.after.title {
        target.title.clone_from(&edit.title_after);
    } else {
        target.title.clone_from(&change.after.title);
    }
    if change.before.author != change.after.author {
        target.author.clone_from(&change.after.author);
    }
    if change.before.subject != change.after.subject {
        target.subject.clone_from(&change.after.subject);
    }
    if change.before.description != change.after.description {
        target.description.clone_from(&change.after.description);
    }
    if change.before.keywords != change.after.keywords {
        target.keywords.clone_from(&change.after.keywords);
    }
    edit.set_metadata(target)?;
    Ok(())
}

struct DurableDirection<'a> {
    source_id: &'a str,
    target_id: &'a str,
    target_bytes: &'a [u8],
}

fn durable_limits() -> PatchLimits {
    PatchLimits::new(
        BlobLimits::new(1, MAX_PACKAGE_BYTES, MAX_PACKAGE_BYTES),
        MAX_WIRE_JSON_BYTES,
        1,
        4,
        4_096,
        16_384,
    )
}

fn durable_operation(
    limits: PatchLimits,
    source: &BlobId,
    target: &BlobId,
) -> Result<PatchOperation> {
    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        SOURCE_PRECONDITION.to_string(),
        Value::String(source.as_hex()),
    );
    PatchOperation::new(
        limits,
        DURABLE_OPERATION,
        DURABLE_TARGET,
        preconditions,
        Value::String(target.as_hex()),
    )
    .map_err(patch_wire_error)
}

fn durable_direction<Mode>(patch: &CorePatch<Mode>) -> Result<DurableDirection<'_>> {
    if patch.format() != DURABLE_FORMAT || patch.operations().len() != 1 {
        return Err(invalid("invalid ODM durable patch vocabulary"));
    }
    let operation = &patch.operations()[0];
    if operation.op != DURABLE_OPERATION
        || operation.target != DURABLE_TARGET
        || operation.preconditions.len() != 1
        || patch.blobs().len() != 1
    {
        return Err(invalid("invalid ODM durable patch vocabulary"));
    }
    let source_id = operation
        .preconditions
        .get(SOURCE_PRECONDITION)
        .and_then(Value::as_str)
        .filter(|value| is_digest(value))
        .ok_or_else(|| invalid("invalid ODM durable patch vocabulary"))?;
    let target_id = operation
        .value
        .as_str()
        .filter(|value| is_digest(value))
        .ok_or_else(|| invalid("invalid ODM durable patch vocabulary"))?;
    let blob_id = patch
        .blobs()
        .ids()
        .next()
        .filter(|id| id.as_hex() == target_id)
        .ok_or_else(|| invalid("invalid ODM durable patch vocabulary"))?;
    let target_bytes = patch
        .blobs()
        .get(blob_id)
        .ok_or_else(|| invalid("invalid ODM durable patch vocabulary"))?;
    Ok(DurableDirection {
        source_id,
        target_id,
        target_bytes,
    })
}

fn validate_reversible(patch: &CorePatch<Reversible>) -> Result<()> {
    let forward = durable_direction(patch)?;
    let inverse = patch.inverse();
    let reverse = durable_direction(&inverse)?;
    if forward.source_id != reverse.target_id || forward.target_id != reverse.source_id {
        return Err(invalid("invalid ODM durable patch vocabulary"));
    }
    Master::from_bytes(copy_bytes(forward.target_bytes)?)?;
    Master::from_bytes(copy_bytes(reverse.target_bytes)?)?;
    Ok(())
}

fn validate_sealed(patch: &CorePatch<ForwardOnly>) -> Result<()> {
    Master::from_bytes(copy_bytes(durable_direction(patch)?.target_bytes)?)?;
    Ok(())
}

fn master_from_target<Mode>(patch: &CorePatch<Mode>) -> Result<Master> {
    Master::from_bytes(copy_bytes(durable_direction(patch)?.target_bytes)?)
}

fn is_digest(value: &str) -> bool {
    value.len() == 64
        && value
            .bytes()
            .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
}

fn copy_bytes(source: &[u8]) -> Result<Vec<u8>> {
    if source.len() > MAX_PACKAGE_BYTES {
        return Err(invalid("ODM durable package exceeds the 256 MiB limit"));
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(source.len())
        .map_err(|allocation| Error::Allocation {
            resource: "ODM durable package",
            source: allocation,
        })?;
    bytes.extend_from_slice(source);
    Ok(bytes)
}

fn patch_wire_error(source: litchi_core::PatchError) -> Error {
    let message = format!("invalid ODM durable patch: {source}");
    drop(source);
    Error::InvalidFormat(message)
}

fn invalid(message: &str) -> Error {
    Error::InvalidFormat(message.to_owned())
}
