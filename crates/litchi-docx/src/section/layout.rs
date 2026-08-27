//! Exact-source transactions for the locally authored layout of one existing
//! main-story section.
//!
//! This capability deliberately sits beside the read-only section inventory.
//! It edits only an existing paragraph-owned or body-final `w:sectPr`; an
//! implicit section and a table-cell `w:sectPr` have no selectable physical
//! owner.  The semantic section codec validates the projected values, while
//! the publication splice retains opaque attributes on edited known nodes and
//! all unrelated section markup.

use super::codec::section_child_rank;
use super::inventory::{Inventory, RelationshipBinding, SourceSpan};
use super::{Columns, Limits, Margins, PageSize, Section, Selector, Start};
use crate::error::{Error, Result};
use crate::namespace::is_wordprocessing_namespace;
use litchi_core::{Position, SourceVersion};
use litchi_opc::{SourceArtifact, SourceArtifactFingerprint, SourceLineage};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;
use std::sync::Arc;

/// An immutable, exact-source section-layout snapshot.
#[derive(Clone)]
pub struct Snapshot {
    xml: Arc<Vec<u8>>,
    inventory: Inventory,
    source_version: Option<SourceVersion>,
    lineage: Option<SourceLineage>,
    artifact_fingerprint: Option<SourceArtifactFingerprint>,
    limits: Limits,
}

impl Snapshot {
    /// Parse an owned main-document XML snapshot.
    pub fn from_xml(xml: Vec<u8>) -> Result<Self> {
        Self::from_xml_with_limits(xml, &Limits::default())
    }

    /// Parse an owned main-document XML snapshot with bounded section limits.
    pub fn from_xml_with_limits(xml: Vec<u8>, limits: &Limits) -> Result<Self> {
        Self::from_shared_xml(Arc::new(xml), None, None, None, limits.clone())
    }

    /// Parse a source-backed main-document payload without detaching its
    /// immutable allocation.
    pub(crate) fn from_source_xml(
        xml: Arc<Vec<u8>>,
        source_version: SourceVersion,
        lineage: SourceLineage,
        artifact_fingerprint: SourceArtifactFingerprint,
        limits: &Limits,
    ) -> Result<Self> {
        Self::from_shared_xml(
            xml,
            Some(source_version),
            Some(lineage),
            Some(artifact_fingerprint),
            limits.clone(),
        )
    }

    fn from_shared_xml(
        xml: Arc<Vec<u8>>,
        source_version: Option<SourceVersion>,
        lineage: Option<SourceLineage>,
        artifact_fingerprint: Option<SourceArtifactFingerprint>,
        limits: Limits,
    ) -> Result<Self> {
        crate::source_backed::ensure_source_section_inventory_xml(&xml, &limits)?;
        let inventory = Inventory::parse_with_limits(xml.as_slice(), &limits)?;
        Ok(Self {
            xml,
            inventory,
            source_version,
            lineage,
            artifact_fingerprint,
            limits,
        })
    }

    /// Borrow the immutable section inventory.
    #[must_use]
    pub const fn inventory(&self) -> &Inventory {
        &self.inventory
    }

    /// Resolve one typed section descriptor.
    #[must_use]
    pub fn section(&self, selector: impl Into<Selector>) -> Option<&super::Descriptor> {
        self.inventory.section(selector)
    }

    /// Start an isolated edit of one existing, locally authored section.
    pub fn edit(&self, selector: impl Into<Selector>) -> Result<Edit> {
        Edit::new(self.clone(), selector.into())
    }

    pub(crate) fn limits(&self) -> &Limits {
        &self.limits
    }

    pub(crate) fn shared_xml(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.xml)
    }

    pub(crate) fn with_artifact_fingerprint(
        &self,
        artifact_fingerprint: SourceArtifactFingerprint,
    ) -> Self {
        Self {
            xml: Arc::clone(&self.xml),
            inventory: self.inventory.clone(),
            source_version: None,
            lineage: None,
            artifact_fingerprint: Some(artifact_fingerprint),
            limits: self.limits.clone(),
        }
    }
}

/// An isolated edit over one existing section's local page-layout properties.
#[derive(Clone)]
pub struct Edit {
    base: Snapshot,
    projected: Section,
    selector: Selector,
    position: Position,
    source_span: SourceSpan,
    source_context: FragmentContext,
    base_state: LocalState,
    projected_state: LocalState,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct LocalState {
    page_size: Option<PageSize>,
    margins: Option<Margins>,
    start: Option<Start>,
    columns: Option<Columns>,
}

impl LocalState {
    fn from_section(section: &mut Section) -> Result<Self> {
        let state = section.local_state()?;
        Ok(Self {
            page_size: state.page_size,
            margins: state.margins,
            start: state.start,
            columns: state.columns,
        })
    }
}

impl Edit {
    fn new(base: Snapshot, selector: Selector) -> Result<Self> {
        let descriptor = base
            .inventory
            .section(selector)
            .ok_or_else(|| match selector {
                Selector::Position(position) => Error::OutOfBounds {
                    object: "DOCX section",
                    index: position.get(),
                    len: base.inventory.sections().len(),
                },
                Selector::Owner(_) => {
                    Error::InvalidFormat("the requested DOCX section owner does not exist".into())
                },
            })?;
        let position = descriptor.position();
        let Some((source_span, _fragment)) = base.inventory.source_fragment(
            base.xml.as_slice(),
            selector,
            base.limits.max_section_bytes,
        )?
        else {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "edit_section_layout",
                reason: "the selected section has no locally authored main-story w:sectPr",
            });
        };
        let source_context =
            FragmentContext::from_source(base.xml.as_slice(), &source_span, &base.limits)?;
        let base_state = LocalState {
            page_size: descriptor.page_size(),
            margins: descriptor.margins(),
            start: descriptor.start(),
            columns: descriptor.columns(),
        };
        let mut projected = Section::default();
        if let Some(value) = base_state.page_size {
            projected.set_page_size(value)?;
        }
        if let Some(value) = base_state.margins {
            projected.set_margins(value)?;
        }
        projected.set_start(base_state.start)?;
        projected.set_columns(base_state.columns.clone())?;
        let base_state = LocalState::from_section(&mut projected)?;
        Ok(Self {
            base,
            projected,
            selector,
            position,
            source_span,
            source_context,
            projected_state: base_state.clone(),
            base_state,
        })
    }

    /// Borrow the exact immutable source snapshot.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.base
    }

    /// Return the checked semantic source-order section position.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Return the projected local page geometry.
    pub fn page_size(&mut self) -> Result<Option<PageSize>> {
        Ok(self.projected_state.page_size)
    }

    /// Return the projected local margins.
    pub fn margins(&mut self) -> Result<Option<Margins>> {
        Ok(self.projected_state.margins)
    }

    /// Return the projected local section-break placement.
    pub fn start(&mut self) -> Result<Option<Start>> {
        Ok(self.projected_state.start)
    }

    /// Return the projected local columns.
    pub fn columns(&mut self) -> Result<Option<Columns>> {
        Ok(self.projected_state.columns.clone())
    }

    /// Whether the projected semantic state differs from the source state.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.projected_state != self.base_state
    }

    /// Replace or remove the locally authored page geometry.
    pub fn set_page_size(&mut self, value: Option<PageSize>) -> Result<&mut Self> {
        if self.projected_state.page_size == value {
            return Ok(self);
        }
        match value {
            Some(value) => self.projected.set_page_size(value)?,
            None => self.projected.clear_page_size()?,
        }
        self.projected_state = LocalState::from_section(&mut self.projected)?;
        Ok(self)
    }

    /// Remove the locally authored page geometry.
    pub fn clear_page_size(&mut self) -> Result<&mut Self> {
        self.set_page_size(None)
    }

    /// Replace or remove the locally authored margins.
    pub fn set_margins(&mut self, value: Option<Margins>) -> Result<&mut Self> {
        if self.projected_state.margins == value {
            return Ok(self);
        }
        match value {
            Some(value) => self.projected.set_margins(value)?,
            None => self.projected.clear_margins()?,
        }
        self.projected_state = LocalState::from_section(&mut self.projected)?;
        Ok(self)
    }

    /// Remove the locally authored margins.
    pub fn clear_margins(&mut self) -> Result<&mut Self> {
        self.set_margins(None)
    }

    /// Replace or remove the locally authored section-break placement.
    pub fn set_start(&mut self, value: Option<Start>) -> Result<&mut Self> {
        if self.projected_state.start == value {
            return Ok(self);
        }
        self.projected.set_start(value)?;
        self.projected_state = LocalState::from_section(&mut self.projected)?;
        Ok(self)
    }

    /// Remove the locally authored section-break placement.
    pub fn clear_start(&mut self) -> Result<&mut Self> {
        self.set_start(None)
    }

    /// Replace or remove the locally authored columns.
    pub fn set_columns(&mut self, value: Option<Columns>) -> Result<&mut Self> {
        if self.projected_state.columns == value {
            return Ok(self);
        }
        self.projected.set_columns(value)?;
        self.projected_state = LocalState::from_section(&mut self.projected)?;
        Ok(self)
    }

    /// Remove the locally authored columns.
    pub fn clear_columns(&mut self) -> Result<&mut Self> {
        self.set_columns(None)
    }

    /// Finish the edit as an exact, reversible patch.
    pub fn commit(mut self) -> Result<Commit> {
        let projected_state = self.projected_state.clone();
        let semantic_noop = projected_state == self.base_state;
        let replacement = if semantic_noop {
            None
        } else {
            let candidate = self.projected.to_xml_bytes()?;
            let (start, end) = self.source_span.range();
            let original =
                self.base.xml.get(start..end).ok_or_else(|| {
                    Error::InvalidFormat("section source span is outside XML".into())
                })?;
            Some(rewrite_known_children(
                original,
                &candidate,
                &self.base_state,
                &projected_state,
                &self.source_context,
            )?)
        };
        if let Some(replacement) = &replacement
            && replacement.len() > self.base.limits.max_section_bytes
        {
            return Err(Error::SectionInventoryLimit {
                resource: "section bytes",
                maximum: self.base.limits.max_section_bytes,
                actual: replacement.len(),
            });
        }
        let target = if semantic_noop {
            self.base.clone()
        } else {
            let (start, end) = self.source_span.range();
            let replacement = replacement.as_deref().ok_or_else(|| {
                Error::InvalidFormat("changed section layout has no replacement bytes".into())
            })?;
            let next_xml = replace_range(self.base.xml.as_slice(), start, end, replacement)?;
            let target = Snapshot::from_shared_xml(
                Arc::new(next_xml),
                self.base.source_version,
                self.base.lineage.clone(),
                self.base.artifact_fingerprint,
                self.base.limits.clone(),
            )?;
            ensure_same_topology(&self.base.inventory, &target.inventory)?;
            ensure_semantic_readback(&target.inventory, self.selector, &projected_state)?;
            target
        };
        if semantic_noop {
            ensure_semantic_readback(&target.inventory, self.selector, &projected_state)?;
        }
        let patch = Patch {
            before: self.base.shared_xml(),
            after: target.shared_xml(),
            authorization: PatchAuthorization::from_snapshot(&self.base)?,
            limits: self.base.limits.clone(),
            selector: self.selector,
        };
        Ok(Commit {
            snapshot: target,
            patch,
        })
    }
}

/// A successful section-layout commit.
#[derive(Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the projected immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the exact reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Move the projected snapshot out of the commit.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

/// An exact-source-checked, reversible section-layout patch.
#[derive(Clone)]
pub struct Patch {
    before: Arc<Vec<u8>>,
    after: Arc<Vec<u8>>,
    authorization: PatchAuthorization,
    limits: Limits,
    selector: Selector,
}

#[derive(Clone)]
enum PatchAuthorization {
    Detached,
    SourceBound {
        source_version: SourceVersion,
        lineage: SourceLineage,
        artifact_fingerprint: SourceArtifactFingerprint,
    },
    EmittedArtifact {
        artifact_fingerprint: SourceArtifactFingerprint,
    },
}

impl PatchAuthorization {
    fn from_snapshot(snapshot: &Snapshot) -> Result<Self> {
        match (
            snapshot.source_version,
            snapshot.lineage.clone(),
            snapshot.artifact_fingerprint,
        ) {
            (None, None, None) => Ok(Self::Detached),
            (Some(source_version), Some(lineage), Some(artifact_fingerprint)) => {
                Ok(Self::SourceBound {
                    source_version,
                    lineage,
                    artifact_fingerprint,
                })
            },
            _ => Err(Error::SectionLayoutAuthorizationConflict),
        }
    }
}

/// Exact publication evidence for one changed section-layout artifact.
///
/// The retained inverse is authorized for the exact ZIP artifact emitted by
/// the forward publication. Use the package inverse-publication method to
/// reauthorize that inverse after reopening the emitted artifact.
pub struct Publication {
    snapshot: Snapshot,
    original_snapshot: Snapshot,
    original_artifact: SourceArtifact,
    inverse_patch: Patch,
}

impl Publication {
    pub(crate) fn new(
        snapshot: Snapshot,
        original_snapshot: Snapshot,
        original_artifact: SourceArtifact,
        inverse_patch: Patch,
    ) -> Self {
        Self {
            snapshot,
            original_snapshot,
            original_artifact,
            inverse_patch,
        }
    }

    /// Borrow the semantic snapshot represented by the emitted artifact.
    ///
    /// Its source identity is the emitted artifact fingerprint retained by
    /// the publication authorization; the process-local source lineage and
    /// revision are intentionally not reused for a reopened artifact.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the inverse patch authorized for the emitted artifact.
    ///
    /// Applying this patch is still source-checked; a reopened package must
    /// be explicitly reauthorized through the package inverse-publication
    /// method rather than by bypassing the complete-artifact check.
    #[must_use]
    pub const fn inverse_patch(&self) -> &Patch {
        &self.inverse_patch
    }

    pub(crate) const fn original_snapshot(&self) -> &Snapshot {
        &self.original_snapshot
    }

    pub(crate) const fn original_artifact(&self) -> &SourceArtifact {
        &self.original_artifact
    }
}

impl Patch {
    /// Whether this patch is an exact byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.as_slice() == self.after.as_slice()
    }

    /// Whether this patch changes the main-document bytes.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.is_noop()
    }

    /// The selector retained as the semantic patch coordinate.
    #[must_use]
    pub const fn selector(&self) -> Selector {
        self.selector
    }

    /// Construct the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            authorization: self.authorization.clone(),
            limits: self.limits.clone(),
            selector: self.selector,
        }
    }

    /// Apply only to a snapshot with the exact source bytes, source lineage,
    /// and semantic limits used to create this patch. The public inverse
    /// retained by a [`Publication`] is the explicit fingerprint-only
    /// reauthorization mode for the emitted artifact.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.xml.as_slice() != self.before.as_slice() {
            return Err(Error::SectionLayoutStaleSource);
        }
        match &self.authorization {
            PatchAuthorization::Detached => {
                if source.source_version.is_some()
                    || source.lineage.is_some()
                    || source.artifact_fingerprint.is_some()
                {
                    return Err(Error::SectionLayoutAuthorizationConflict);
                }
            },
            PatchAuthorization::SourceBound {
                source_version,
                lineage,
                artifact_fingerprint,
            } => {
                if source.lineage.as_ref() != Some(lineage) {
                    return Err(Error::SectionLayoutForeignSource);
                }
                if source.source_version != Some(*source_version) {
                    return Err(Error::SectionLayoutRevisionConflict);
                }
                if source.artifact_fingerprint != Some(*artifact_fingerprint) {
                    return Err(Error::SectionLayoutFingerprintConflict);
                }
            },
            PatchAuthorization::EmittedArtifact {
                artifact_fingerprint,
            } => {
                if source.artifact_fingerprint != Some(*artifact_fingerprint) {
                    return Err(Error::SectionLayoutFingerprintConflict);
                }
            },
        }
        if source.limits != self.limits {
            return Err(Error::SectionLayoutPolicyConflict);
        }
        if self.is_noop() {
            return Ok(source.clone());
        }
        Snapshot::from_shared_xml(
            Arc::clone(&self.after),
            source.source_version,
            source.lineage.clone(),
            source.artifact_fingerprint,
            self.limits.clone(),
        )
    }

    pub(crate) fn limits(&self) -> &Limits {
        &self.limits
    }

    pub(crate) fn reauthorize_for_artifact(&mut self, fingerprint: SourceArtifactFingerprint) {
        self.authorization = PatchAuthorization::EmittedArtifact {
            artifact_fingerprint: fingerprint,
        };
    }
}

#[derive(Debug, Clone)]
struct ChildSpan {
    start: usize,
    end: usize,
    local_start: usize,
    local_end: usize,
    word: bool,
}

impl ChildSpan {
    fn local<'a>(&self, xml: &'a [u8]) -> &'a [u8] {
        &xml[self.local_start..self.local_end]
    }
}

#[derive(Clone)]
struct FragmentContext {
    analysis: FragmentAnalysis,
}

#[derive(Clone)]
struct FragmentAnalysis {
    elements: Vec<ElementInfo>,
}

#[derive(Clone)]
struct ElementInfo {
    start: usize,
    end: usize,
    local_start: usize,
    local_end: usize,
    word: bool,
    word_namespace: Option<Vec<u8>>,
    namespace_bindings: Vec<RelationshipBinding>,
    attributes: WordAttributeInfo,
    parent: Option<usize>,
}

impl FragmentContext {
    fn from_source(source: &[u8], span: &SourceSpan, limits: &Limits) -> Result<Self> {
        let (start, end) = span.range();
        let section = source.get(start..end).ok_or_else(|| {
            Error::InvalidFormat("section source span is outside document XML".into())
        })?;
        let opening_end = opening_tag_end(section)?;
        let root_attributes = lexical_attributes(&section[..opening_end])?;
        let mut wrapper_prefix = b"__litchi_section_layout".to_vec();
        let mut suffix = 0usize;
        while span
            .namespace_bindings()
            .iter()
            .any(|binding| binding.prefix == wrapper_prefix)
            || root_attributes
                .iter()
                .any(|attribute| attribute.prefix == b"xmlns" && attribute.local == wrapper_prefix)
        {
            suffix = suffix
                .checked_add(1)
                .ok_or_else(|| Error::InvalidFormat("section wrapper prefix overflow".into()))?;
            wrapper_prefix = format!("__litchi_section_layout{suffix}").into_bytes();
        }

        let mut prefix = Vec::new();
        prefix.extend_from_slice(b"<");
        prefix.extend_from_slice(&wrapper_prefix);
        prefix.extend_from_slice(b":root xmlns:");
        prefix.extend_from_slice(&wrapper_prefix);
        prefix.extend_from_slice(b"=\"urn:litchi:section-layout\"");
        for binding in span.namespace_bindings() {
            if binding.prefix.is_empty() {
                prefix.extend_from_slice(b" xmlns=\"");
            } else {
                prefix.extend_from_slice(b" xmlns:");
                prefix.extend_from_slice(&binding.prefix);
                prefix.extend_from_slice(b"=\"");
            }
            append_xml_attribute_value(&mut prefix, &binding.namespace, b'"');
            prefix.push(b'"');
        }
        prefix.extend_from_slice(b">");
        let section_start = prefix.len();
        let suffix_bytes = {
            let mut bytes = Vec::new();
            bytes.extend_from_slice(b"</");
            bytes.extend_from_slice(&wrapper_prefix);
            bytes.extend_from_slice(b":root>");
            bytes
        };
        let capacity = prefix
            .len()
            .checked_add(section.len())
            .and_then(|value| value.checked_add(suffix_bytes.len()))
            .ok_or_else(|| Error::InvalidFormat("section wrapper size overflow".into()))?;
        enforce_layout_limit("section bytes", section.len(), limits.max_section_bytes)?;
        let wrapper_limit = limits
            .max_section_bytes
            .checked_add(prefix.len())
            .and_then(|value| value.checked_add(suffix_bytes.len()))
            .ok_or_else(|| Error::InvalidFormat("section wrapper limit overflow".into()))?;
        enforce_layout_limit("section bytes", capacity, wrapper_limit)?;
        let mut xml = Vec::new();
        xml.try_reserve_exact(capacity)
            .map_err(|source| Error::Allocation {
                resource: "section layout fragment wrapper",
                source,
            })?;
        xml.extend_from_slice(&prefix);
        xml.extend_from_slice(section);
        xml.extend_from_slice(&suffix_bytes);
        let analysis = FragmentAnalysis::parse(&xml, section_start, section.len(), limits)?;
        Ok(Self { analysis })
    }

    fn element(&self, start: usize) -> Result<&ElementInfo> {
        self.analysis
            .elements
            .iter()
            .find(|element| element.start == start)
            .ok_or_else(|| {
                Error::InvalidFormat("section layout source element was not indexed".into())
            })
    }

    fn children(&self, parent_start: usize) -> Result<Vec<ChildSpan>> {
        let parent = self
            .analysis
            .elements
            .iter()
            .position(|element| element.start == parent_start)
            .ok_or_else(|| {
                Error::InvalidFormat("section layout parent element was not indexed".into())
            })?;
        let mut children = Vec::new();
        for element in &self.analysis.elements {
            if element.parent != Some(parent) {
                continue;
            }
            let start = element.start.checked_sub(parent_start).ok_or_else(|| {
                Error::InvalidFormat("section child offset precedes its parent".into())
            })?;
            let end = element.end.checked_sub(parent_start).ok_or_else(|| {
                Error::InvalidFormat("section child end precedes its parent".into())
            })?;
            let local_start = element
                .local_start
                .checked_sub(parent_start)
                .ok_or_else(|| {
                    Error::InvalidFormat("section child local offset precedes its parent".into())
                })?;
            let local_end = element.local_end.checked_sub(parent_start).ok_or_else(|| {
                Error::InvalidFormat("section child local end precedes its parent".into())
            })?;
            children.push(ChildSpan {
                start,
                end,
                local_start,
                local_end,
                word: element.word,
            });
        }
        Ok(children)
    }

    fn in_scope_word_attribute_prefix(
        &self,
        target_start: usize,
        preferred: Option<&[u8]>,
    ) -> Result<Option<(Vec<u8>, Vec<u8>)>> {
        let element = self.element(target_start)?;
        let Some(owner) = element.word_namespace.as_deref() else {
            return Ok(None);
        };
        Ok(element
            .namespace_bindings
            .iter()
            .rev()
            .find(|binding| {
                !binding.prefix.is_empty()
                    && binding.namespace.as_slice() == owner
                    && preferred.is_none_or(|wanted| wanted == binding.prefix.as_slice())
            })
            .map(|binding| (binding.prefix.clone(), binding.namespace.clone())))
    }

    fn fresh_word_attribute_prefix(&self, target_start: usize) -> Result<(Vec<u8>, Vec<u8>)> {
        let element = self.element(target_start)?;
        let namespace = element.word_namespace.clone().ok_or(Error::UnsafeEdit {
            format: "DOCX",
            operation: "edit_section_layout",
            reason: "the target section element has no resolved owning Word namespace",
        })?;
        let mut suffix = 0usize;
        loop {
            let prefix = if suffix == 0 {
                b"w".to_vec()
            } else {
                format!("w{suffix}").into_bytes()
            };
            if !element
                .namespace_bindings
                .iter()
                .any(|binding| binding.prefix == prefix)
            {
                return Ok((prefix, namespace));
            }
            suffix = suffix.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("section layout namespace prefix counter overflow".into())
            })?;
        }
    }
}

impl FragmentAnalysis {
    fn parse(
        xml: &[u8],
        section_start: usize,
        section_len: usize,
        limits: &Limits,
    ) -> Result<Self> {
        let section_end = section_len;
        let wrapper_section_end = section_start
            .checked_add(section_len)
            .ok_or_else(|| Error::InvalidFormat("section fragment range overflow".into()))?;
        if wrapper_section_end > xml.len() {
            return Err(Error::InvalidFormat(
                "section fragment range is outside its wrapper".into(),
            ));
        }
        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().trim_text(false);
        let mut depth = 0usize;
        let mut events = 0usize;
        let mut stack = Vec::<usize>::new();
        let mut elements = Vec::<ElementInfo>::new();
        let mut section_root = None::<usize>;
        let mut section_closed = false;

        loop {
            let start = reader_offset(&reader)?;
            let event = reader
                .read_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            let end = reader_offset(&reader)?;
            events = events
                .checked_add(1)
                .ok_or_else(|| Error::InvalidFormat("section event counter overflow".into()))?;
            enforce_layout_limit("XML events", events, limits.max_events)?;
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) => {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| Error::InvalidFormat("section XML depth overflow".into()))?;
                    let relative_depth = depth.saturating_sub(1);
                    enforce_layout_limit("XML depth", relative_depth, limits.max_depth)?;
                    let is_section_root = start == section_start;
                    if is_section_root {
                        if section_root.is_some()
                            || !is_wordprocessing_namespace(&namespace)
                            || element.local_name().as_ref() != b"sectPr"
                        {
                            return Err(Error::InvalidFormat(
                                "section fragment root is not WordprocessingML sectPr".into(),
                            ));
                        }
                    }
                    let in_section = section_root.is_some() || is_section_root;
                    if in_section {
                        let index = push_fragment_element(
                            &mut elements,
                            start,
                            0,
                            &element,
                            &namespace,
                            &resolver,
                            stack.last().copied(),
                            section_start,
                        )?;
                        if is_section_root {
                            section_root = Some(index);
                        }
                        stack.push(index);
                    }
                },
                Event::Empty(element) => {
                    let child_depth = depth
                        .checked_add(1)
                        .ok_or_else(|| Error::InvalidFormat("section XML depth overflow".into()))?;
                    let relative_depth = child_depth.saturating_sub(1);
                    enforce_layout_limit("XML depth", relative_depth, limits.max_depth)?;
                    let is_section_root = start == section_start;
                    if is_section_root {
                        if section_root.is_some()
                            || !is_wordprocessing_namespace(&namespace)
                            || element.local_name().as_ref() != b"sectPr"
                        {
                            return Err(Error::InvalidFormat(
                                "section fragment root is not WordprocessingML sectPr".into(),
                            ));
                        }
                    }
                    if section_root.is_some() || is_section_root {
                        let index = push_fragment_element(
                            &mut elements,
                            start,
                            end,
                            &element,
                            &namespace,
                            &resolver,
                            stack.last().copied(),
                            section_start,
                        )?;
                        if is_section_root {
                            section_root = Some(index);
                            section_closed = true;
                        }
                    }
                },
                Event::End(_) => {
                    if let Some(index) = stack.pop() {
                        elements[index].end = end.checked_sub(section_start).ok_or_else(|| {
                            Error::InvalidFormat("section element end precedes source range".into())
                        })?;
                        if Some(index) == section_root {
                            section_closed = true;
                        }
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("section XML depth underflow".into())
                    })?;
                },
                Event::Eof => break,
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => {},
            }
        }

        let root = section_root
            .ok_or_else(|| Error::InvalidFormat("section fragment root was not found".into()))?;
        if !section_closed || !stack.is_empty() || elements[root].end != section_end {
            return Err(Error::InvalidFormat(
                "section fragment root range does not match its source span".into(),
            ));
        }
        Ok(Self { elements })
    }
}

fn push_fragment_element(
    elements: &mut Vec<ElementInfo>,
    start: usize,
    end: usize,
    element: &BytesStart<'_>,
    namespace: &ResolveResult<'_>,
    resolver: &NamespaceResolver,
    parent: Option<usize>,
    section_start: usize,
) -> Result<usize> {
    let qualified_name = element.name();
    let name = qualified_name.as_ref();
    let local_offset = name
        .iter()
        .position(|byte| *byte == b':')
        .map_or(0, |index| index + 1);
    let raw_start = start.checked_sub(section_start).ok_or_else(|| {
        Error::InvalidFormat("section element offset precedes source range".into())
    })?;
    let local_start = raw_start
        .checked_add(1)
        .and_then(|value| value.checked_add(local_offset))
        .ok_or_else(|| Error::InvalidFormat("section element local offset overflow".into()))?;
    let local_end = raw_start
        .checked_add(1)
        .and_then(|value| value.checked_add(name.len()))
        .ok_or_else(|| Error::InvalidFormat("section element local range overflow".into()))?;
    let end = if end == 0 {
        0
    } else {
        end.checked_sub(section_start).ok_or_else(|| {
            Error::InvalidFormat("section element end precedes source range".into())
        })?
    };
    let index = elements.len();
    elements.push(ElementInfo {
        start: raw_start,
        end,
        local_start,
        local_end,
        word: is_wordprocessing_namespace(namespace),
        word_namespace: match namespace {
            ResolveResult::Bound(Namespace(namespace)) if is_word_namespace_bytes(namespace) => {
                Some(namespace.to_vec())
            },
            _ => None,
        },
        namespace_bindings: effective_namespace_bindings(resolver),
        attributes: word_attribute_info_from_element(element, namespace, resolver)?,
        parent,
    });
    Ok(index)
}

fn effective_namespace_bindings(resolver: &NamespaceResolver) -> Vec<RelationshipBinding> {
    resolver
        .bindings()
        .map(|(prefix, Namespace(namespace))| RelationshipBinding {
            prefix: match prefix {
                PrefixDeclaration::Default => Vec::new(),
                PrefixDeclaration::Named(prefix) => prefix.as_ref().to_vec(),
            },
            namespace: namespace.to_vec(),
        })
        .collect()
}

fn word_attribute_info_from_element(
    element: &BytesStart<'_>,
    namespace: &ResolveResult<'_>,
    resolver: &NamespaceResolver,
) -> Result<WordAttributeInfo> {
    if !is_wordprocessing_namespace(namespace) {
        return Ok(WordAttributeInfo::default());
    }
    let mut info = WordAttributeInfo::default();
    let element_is_unprefixed = element.name().prefix().is_none();
    let element_namespace = match namespace {
        ResolveResult::Bound(Namespace(namespace)) if is_word_namespace_bytes(namespace) => {
            Some(*namespace)
        },
        _ => None,
    };
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let local = attribute.key.local_name();
        let prefix = attribute
            .key
            .prefix()
            .map_or_else(Vec::new, |prefix| prefix.as_ref().to_vec());
        if is_namespace_attribute(&prefix, local.as_ref()) {
            continue;
        }
        let (attribute_namespace, _) = resolver.resolve_attribute(attribute.key);
        let word = matches!(
            (&element_namespace, &attribute_namespace),
            (
                Some(element),
                ResolveResult::Bound(Namespace(attribute)),
            ) if element == attribute
        ) || (element_is_unprefixed
            && prefix.is_empty()
            && matches!(attribute_namespace, ResolveResult::Unbound));
        if word {
            if info.any_prefix.is_none() {
                info.any_prefix = Some(prefix.clone());
            }
            info.prefixes.push((local.as_ref().to_vec(), prefix));
        }
    }
    Ok(info)
}

fn enforce_layout_limit(resource: &'static str, actual: usize, maximum: usize) -> Result<()> {
    if actual > maximum {
        return Err(Error::SectionInventoryLimit {
            resource,
            maximum,
            actual,
        });
    }
    Ok(())
}

fn append_xml_attribute_value(output: &mut Vec<u8>, value: &[u8], quote: u8) {
    for byte in value {
        match *byte {
            b'&' => output.extend_from_slice(b"&amp;"),
            b'<' => output.extend_from_slice(b"&lt;"),
            b'>' => output.extend_from_slice(b"&gt;"),
            b'"' if quote == b'"' => output.extend_from_slice(b"&quot;"),
            b'\'' if quote == b'\'' => output.extend_from_slice(b"&apos;"),
            byte => output.push(byte),
        }
    }
}

fn rewrite_known_children(
    original: &[u8],
    candidate: &[u8],
    before: &LocalState,
    after: &LocalState,
    fragment: &FragmentContext,
) -> Result<Vec<u8>> {
    if is_self_closing_root(original)
        && [
            KnownChild::Type,
            KnownChild::PageSize,
            KnownChild::Margins,
            KnownChild::Columns,
        ]
        .iter()
        .any(|field| field.changed(before, after) && field.is_present(after))
    {
        let expanded = expand_self_closing_root(original)?;
        return rewrite_known_children(&expanded, candidate, before, after, fragment);
    }
    let original_children = direct_children_with_context(original, fragment, 0)?;
    let candidate_children = direct_children(candidate)?;
    let fields = [
        KnownChild::Type,
        KnownChild::PageSize,
        KnownChild::Margins,
        KnownChild::Columns,
    ];
    let mut replacements = Vec::new();
    for field in fields {
        if !field.changed(before, after) {
            continue;
        }
        let source = unique_child(&original_children, original, field.local())?;
        let target = unique_child(&candidate_children, candidate, field.local())?;
        match (source, target, field.is_present(after)) {
            (Some(source), Some(target), true) => {
                let source_bytes = &original[source.start..source.end];
                let target_bytes = &candidate[target.start..target.end];
                let replacement = match field {
                    KnownChild::Columns => {
                        patch_columns(fragment, source.start, source_bytes, target_bytes)?
                    },
                    _ => patch_known_attributes_at(
                        fragment,
                        source.start,
                        source_bytes,
                        target_bytes,
                        field.attributes(),
                    )?,
                };
                replacements.push((source.start, source.end, replacement));
            },
            (Some(source), None, false) => {
                ensure_clear_is_lossless(
                    fragment,
                    source.start,
                    &original[source.start..source.end],
                    field,
                )?;
                replacements.push((source.start, source.end, Vec::new()));
            },
            (None, Some(target), true) => {
                let replacement = rebind_candidate_child(
                    &candidate[target.start..target.end],
                    original,
                    fragment,
                )?;
                let insertion = insertion_offset(original, &original_children, field)?;
                replacements.push((insertion, insertion, replacement));
            },
            (None, None, false) => {},
            _ => {
                return Err(Error::UnsafeEdit {
                    format: "DOCX",
                    operation: "edit_section_layout",
                    reason: "the section layout codec could not express the requested local property",
                });
            },
        }
    }
    replacements.sort_by_key(|(start, _, _)| *start);
    let mut output = original.to_vec();
    for (start, end, replacement) in replacements.into_iter().rev() {
        output = replace_range(&output, start, end, &replacement)?;
    }
    Ok(output)
}

fn ensure_clear_is_lossless(
    fragment: &FragmentContext,
    target_start: usize,
    original: &[u8],
    field: KnownChild,
) -> Result<()> {
    let opening_end = opening_tag_end(original)?;
    let attributes = lexical_attributes(&original[..opening_end])?;
    let info = word_attribute_info(fragment, target_start)?;
    if field.attributes().iter().any(|local| {
        attributes
            .iter()
            .filter(|attribute| attribute.local.as_slice() == *local)
            .count()
            > 1
    }) {
        return Err(Error::UnsafeEdit {
            format: "DOCX",
            operation: "edit_section_layout",
            reason: "clearing the section property would discard duplicate markup",
        });
    }
    if attributes.iter().any(|attribute| {
        attribute.prefix == b"xmlns"
            || (attribute.prefix.is_empty() && attribute.local.as_slice() == b"xmlns")
            || !field.attributes().contains(&attribute.local.as_slice())
            || !info.prefixes.iter().any(|(local, prefix)| {
                local.as_slice() == attribute.local.as_slice() && prefix == &attribute.prefix
            })
    }) {
        return Err(Error::UnsafeEdit {
            format: "DOCX",
            operation: "edit_section_layout",
            reason: "clearing the section property would discard opaque markup",
        });
    }
    if !is_self_closing_root(original) && root_close_start(original)? != opening_end {
        return Err(Error::UnsafeEdit {
            format: "DOCX",
            operation: "edit_section_layout",
            reason: "clearing the section property would discard opaque child content",
        });
    }
    Ok(())
}

#[derive(Debug, Clone, Copy)]
enum KnownChild {
    Type,
    PageSize,
    Margins,
    Columns,
}

impl KnownChild {
    const fn local(self) -> &'static [u8] {
        match self {
            Self::Type => b"type",
            Self::PageSize => b"pgSz",
            Self::Margins => b"pgMar",
            Self::Columns => b"cols",
        }
    }

    const fn rank(self) -> usize {
        section_child_rank(self.local()).expect("modeled section child has a rank") as usize
    }

    const fn attributes(self) -> &'static [&'static [u8]] {
        match self {
            Self::Type => &[b"val"],
            Self::PageSize => &[b"w", b"h", b"orient"],
            Self::Margins => &[
                b"top", b"right", b"bottom", b"left", b"header", b"footer", b"gutter",
            ],
            Self::Columns => &[b"equalWidth", b"num", b"space", b"sep"],
        }
    }

    fn changed(self, before: &LocalState, after: &LocalState) -> bool {
        match self {
            Self::Type => before.start != after.start,
            Self::PageSize => before.page_size != after.page_size,
            Self::Margins => before.margins != after.margins,
            Self::Columns => before.columns != after.columns,
        }
    }

    fn is_present(self, state: &LocalState) -> bool {
        match self {
            Self::Type => state.start.is_some(),
            Self::PageSize => state.page_size.is_some(),
            Self::Margins => state.margins.is_some(),
            Self::Columns => state.columns.is_some(),
        }
    }
}

fn unique_child<'a>(
    children: &'a [ChildSpan],
    xml: &[u8],
    local: &[u8],
) -> Result<Option<&'a ChildSpan>> {
    let mut found = None;
    for child in children {
        if child.word && child.local(xml) == local {
            if found.is_some() {
                return Err(Error::UnsafeEdit {
                    format: "DOCX",
                    operation: "edit_section_layout",
                    reason: "duplicate Word section-layout children are not safely editable",
                });
            }
            found = Some(child);
        }
    }
    Ok(found)
}

fn insertion_offset(original: &[u8], children: &[ChildSpan], field: KnownChild) -> Result<usize> {
    if children
        .iter()
        .any(|child| child.word && section_child_rank(child.local(original)).is_none())
    {
        return Err(Error::UnsafeEdit {
            format: "DOCX",
            operation: "edit_section_layout",
            reason: "an unknown direct Word section child blocks safe modeled insertion",
        });
    }
    for child in children {
        if !child.word {
            continue;
        }
        let Some(other_rank) = section_child_rank(child.local(original)) else {
            continue;
        };
        if usize::from(other_rank) > field.rank() {
            return Ok(child.start);
        }
    }
    root_close_start(original)
}

#[derive(Debug, Clone)]
struct LexicalAttribute {
    start: usize,
    end: usize,
    value_start: usize,
    value_end: usize,
    local: Vec<u8>,
    prefix: Vec<u8>,
    quote: u8,
}

#[derive(Debug, Clone, Default)]
struct WordAttributeInfo {
    prefixes: Vec<(Vec<u8>, Vec<u8>)>,
    any_prefix: Option<Vec<u8>>,
}

fn patch_columns(
    fragment: &FragmentContext,
    target_start: usize,
    original: &[u8],
    candidate: &[u8],
) -> Result<Vec<u8>> {
    let original_children = direct_children_with_context(original, fragment, target_start)?;
    let candidate_children = direct_children(candidate)?;
    let original_columns: Vec<&ChildSpan> = original_children
        .iter()
        .filter(|child| child.word && child.local(original) == b"col")
        .collect();
    let candidate_columns: Vec<&ChildSpan> = candidate_children
        .iter()
        .filter(|child| child.word && child.local(candidate) == b"col")
        .collect();
    if original_children
        .iter()
        .any(|child| child.word && child.local(original) != b"col")
        || original_columns.len() != candidate_columns.len()
    {
        return Err(Error::UnsafeEdit {
            format: "DOCX",
            operation: "edit_section_layout",
            reason: "column child structure changes are outside the lossless range-local closure",
        });
    }

    let mut output = original.to_vec();
    for (source, target) in original_columns.iter().zip(candidate_columns.iter()).rev() {
        let source_bytes = &original[source.start..source.end];
        let target_bytes = &candidate[target.start..target.end];
        let replacement = patch_known_attributes_at(
            fragment,
            target_start.checked_add(source.start).ok_or_else(|| {
                Error::InvalidFormat("section column source offset overflow".into())
            })?,
            source_bytes,
            target_bytes,
            &[b"w", b"space"],
        )?;
        output = replace_range(&output, source.start, source.end, &replacement)?;
    }
    let root_info = word_attribute_info(fragment, target_start)?;
    patch_known_attributes_with_info(
        fragment,
        target_start,
        &output,
        candidate,
        KnownChild::Columns.attributes(),
        &root_info,
    )
}

fn patch_known_attributes_at(
    fragment: &FragmentContext,
    target_start: usize,
    original: &[u8],
    candidate: &[u8],
    known: &[&[u8]],
) -> Result<Vec<u8>> {
    let info = word_attribute_info(fragment, target_start)?;
    patch_known_attributes_with_info(fragment, target_start, original, candidate, known, &info)
}

fn patch_known_attributes_with_info(
    fragment: &FragmentContext,
    target_start: usize,
    original: &[u8],
    candidate: &[u8],
    known: &[&[u8]],
    info: &WordAttributeInfo,
) -> Result<Vec<u8>> {
    let original_end = opening_tag_end(original)?;
    let candidate_end = opening_tag_end(candidate)?;
    let original_attributes = lexical_attributes(&original[..original_end])?;
    let candidate_attributes = lexical_attributes(&candidate[..candidate_end])?;
    let mut edits = Vec::new();
    let mut additions = Vec::new();
    let mut bind_attribute_namespace = None;
    for local in known {
        let desired = candidate_attributes
            .iter()
            .find(|attribute| attribute.local.as_slice() == *local);
        let source_matches: Vec<&LexicalAttribute> = original_attributes
            .iter()
            .filter(|attribute| {
                attribute.local.as_slice() == *local
                    && info.prefixes.iter().any(|(known_local, prefix)| {
                        known_local.as_slice() == *local && prefix == &attribute.prefix
                    })
            })
            .collect();
        if source_matches.len() > 1 {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "edit_section_layout",
                reason: "duplicate expanded Word attributes are not safely editable",
            });
        }
        match (source_matches.first().copied(), desired) {
            (Some(source), Some(target)) => {
                edits.push((
                    source.value_start,
                    source.value_end,
                    candidate[target.value_start..target.value_end].to_vec(),
                ));
            },
            (Some(source), None) => {
                let mut start = source.start;
                while start > 0 && original[start - 1].is_ascii_whitespace() {
                    start -= 1;
                }
                edits.push((start, source.end, Vec::new()));
            },
            (None, Some(target)) => {
                let in_scope_prefix = fragment
                    .in_scope_word_attribute_prefix(target_start, None)?
                    .map(|(prefix, _)| prefix);
                let prefix = info
                    .prefixes
                    .iter()
                    .find(|(known_local, prefix)| {
                        known_local.as_slice() == *local && !prefix.is_empty()
                    })
                    .map(|(_, prefix)| prefix.clone())
                    .or_else(|| {
                        info.any_prefix
                            .as_ref()
                            .filter(|prefix| !prefix.is_empty())
                            .cloned()
                    })
                    .or_else(|| {
                        let element_prefix = element_prefix(&original[..original_end]);
                        (!element_prefix.is_empty()).then(|| element_prefix.to_vec())
                    })
                    .or_else(|| in_scope_prefix.clone());
                let prefix = match prefix {
                    Some(prefix) => prefix,
                    None => {
                        let (prefix, namespace) =
                            fragment.fresh_word_attribute_prefix(target_start)?;
                        bind_attribute_namespace = Some((prefix.clone(), namespace));
                        prefix
                    },
                };
                let mut attribute = Vec::new();
                attribute.push(b' ');
                if !prefix.is_empty() {
                    attribute.extend_from_slice(&prefix);
                    attribute.push(b':');
                }
                attribute.extend_from_slice(local);
                attribute.extend_from_slice(b"=");
                let quote = original_attributes
                    .first()
                    .map_or(b'"', |attribute| attribute.quote);
                attribute.push(quote);
                attribute.extend_from_slice(&candidate[target.value_start..target.value_end]);
                attribute.push(quote);
                additions.push(attribute);
            },
            (None, None) => {},
        }
    }
    if !additions.is_empty() {
        let insertion = if original[..original_end].ends_with(b"/>") {
            original_end - 2
        } else {
            original_end - 1
        };
        let mut addition = Vec::new();
        for attribute in additions {
            addition.extend_from_slice(&attribute);
        }
        edits.push((insertion, insertion, addition));
    }
    edits.sort_by_key(|(start, _, _)| *start);
    let mut output = original.to_vec();
    for (start, end, replacement) in edits.into_iter().rev() {
        output = replace_range(&output, start, end, &replacement)?;
    }
    if let Some((prefix, namespace)) = bind_attribute_namespace {
        output = bind_candidate_attribute_namespace(&output, &prefix, &namespace)?;
    }
    Ok(output)
}

fn word_attribute_info(
    fragment: &FragmentContext,
    target_start: usize,
) -> Result<WordAttributeInfo> {
    let element = fragment.element(target_start)?;
    if !element.word {
        return Err(Error::UnsafeEdit {
            format: "DOCX",
            operation: "edit_section_layout",
            reason: "the target section child is not WordprocessingML",
        });
    }
    Ok(element.attributes.clone())
}

fn lexical_attributes(opening: &[u8]) -> Result<Vec<LexicalAttribute>> {
    let mut attributes = Vec::new();
    let mut cursor = 1usize;
    while cursor < opening.len()
        && !opening[cursor].is_ascii_whitespace()
        && !matches!(opening[cursor], b'>' | b'/')
    {
        cursor += 1;
    }
    while cursor < opening.len() {
        while cursor < opening.len()
            && (opening[cursor].is_ascii_whitespace() || opening[cursor] == b'/')
        {
            cursor += 1;
        }
        if cursor >= opening.len() || opening[cursor] == b'>' {
            break;
        }
        let start = cursor;
        while cursor < opening.len()
            && !opening[cursor].is_ascii_whitespace()
            && !matches!(opening[cursor], b'=' | b'>' | b'/')
        {
            cursor += 1;
        }
        let name = &opening[start..cursor];
        let (prefix, local) = name
            .iter()
            .position(|byte| *byte == b':')
            .map_or((Vec::new(), name.to_vec()), |index| {
                (name[..index].to_vec(), name[index + 1..].to_vec())
            });
        while cursor < opening.len() && opening[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if opening.get(cursor) != Some(&b'=') {
            return Err(Error::InvalidFormat(
                "section attribute has no value".into(),
            ));
        }
        cursor += 1;
        while cursor < opening.len() && opening[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *opening
            .get(cursor)
            .ok_or_else(|| Error::InvalidFormat("section attribute value is incomplete".into()))?;
        if !matches!(quote, b'\'' | b'"') {
            return Err(Error::InvalidFormat(
                "section attribute value is not quoted".into(),
            ));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < opening.len() && opening[cursor] != quote {
            cursor += 1;
        }
        if cursor >= opening.len() {
            return Err(Error::InvalidFormat(
                "section attribute value is unterminated".into(),
            ));
        }
        let value_end = cursor;
        cursor += 1;
        attributes.push(LexicalAttribute {
            start,
            end: cursor,
            value_start,
            value_end,
            local,
            prefix,
            quote,
        });
    }
    Ok(attributes)
}

fn reader_offset(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_error| Error::InvalidFormat("section XML offset overflow".into()))
}

fn raw_reader_offset(reader: &Reader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_error| Error::InvalidFormat("section XML offset overflow".into()))
}

fn has_empty_default_namespace(opening: &[u8]) -> Result<bool> {
    let end = opening_tag_end(opening)?;
    Ok(lexical_attributes(&opening[..end])?
        .iter()
        .any(|attribute| {
            attribute.prefix.is_empty()
                && attribute.local.as_slice() == b"xmlns"
                && attribute.value_start == attribute.value_end
        }))
}

fn direct_children(xml: &[u8]) -> Result<Vec<ChildSpan>> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let root_end = opening_tag_end(xml)?;
    let fragment_prefix = element_prefix(&xml[..root_end]);
    let mut depth = 0usize;
    let mut children = Vec::new();
    loop {
        let start = raw_reader_offset(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = raw_reader_offset(&reader)?;
        match event {
            Event::Start(element) => {
                if depth == 1 {
                    let qualified_name = element.name();
                    let name = qualified_name.as_ref();
                    let child_prefix = name
                        .iter()
                        .position(|byte| *byte == b':')
                        .map_or(name, |index| &name[..index]);
                    let resets_default_namespace = has_empty_default_namespace(&xml[start..end])?;
                    let local_start = name
                        .iter()
                        .position(|byte| *byte == b':')
                        .map_or(0, |index| index + 1);
                    children.push(ChildSpan {
                        start,
                        end: 0,
                        local_start: start + 1 + local_start,
                        local_end: start + 1 + name.len(),
                        word: child_prefix == fragment_prefix
                            && !(fragment_prefix.is_empty() && resets_default_namespace),
                    });
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("section XML depth overflow".into()))?;
            },
            Event::Empty(element) => {
                if depth == 1 {
                    let qualified_name = element.name();
                    let name = qualified_name.as_ref();
                    let child_prefix = name
                        .iter()
                        .position(|byte| *byte == b':')
                        .map_or(name, |index| &name[..index]);
                    let resets_default_namespace = has_empty_default_namespace(&xml[start..end])?;
                    let local_start = name
                        .iter()
                        .position(|byte| *byte == b':')
                        .map_or(0, |index| index + 1);
                    children.push(ChildSpan {
                        start,
                        end,
                        local_start: start + 1 + local_start,
                        local_end: start + 1 + name.len(),
                        word: child_prefix == fragment_prefix
                            && !(fragment_prefix.is_empty() && resets_default_namespace),
                    });
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("section XML depth underflow".into()))?;
                if depth == 1
                    && let Some(child) = children.last_mut()
                    && child.end == 0
                {
                    child.end = end;
                }
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if depth != 0 {
        return Err(Error::InvalidFormat(
            "section XML has unclosed direct child".into(),
        ));
    }
    Ok(children)
}

fn direct_children_with_context(
    xml: &[u8],
    fragment: &FragmentContext,
    target_start: usize,
) -> Result<Vec<ChildSpan>> {
    let parent = fragment.element(target_start)?;
    let expected_len = parent.end.checked_sub(target_start).ok_or_else(|| {
        Error::InvalidFormat("section parent range precedes its source offset".into())
    })?;
    let children = fragment.children(target_start)?;
    if xml.len() != expected_len && !children.is_empty() {
        return Err(Error::InvalidFormat(
            "section source fragment length changed before child lookup".into(),
        ));
    }
    Ok(children)
}

fn root_close_start(xml: &[u8]) -> Result<usize> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    loop {
        let start = raw_reader_offset(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = raw_reader_offset(&reader)?;
        match event {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("section XML depth overflow".into()))?
            },
            Event::Empty(_) if depth == 0 => return Ok(end),
            Event::Empty(_) => {},
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("section XML depth underflow".into()))?;
                if depth == 0 {
                    return Ok(start);
                }
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Err(Error::InvalidFormat(
        "section XML has no closing root element".into(),
    ))
}

fn is_self_closing_root(xml: &[u8]) -> bool {
    opening_tag_end(xml)
        .ok()
        .is_some_and(|end| xml[..end].ends_with(b"/>"))
}

fn expand_self_closing_root(xml: &[u8]) -> Result<Vec<u8>> {
    let opening_end = opening_tag_end(xml)?;
    if !xml[..opening_end].ends_with(b"/>") {
        return Ok(xml.to_vec());
    }
    let name_end = xml[1..]
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(*byte, b'>' | b'/'))
        .map_or(opening_end - 1, |offset| offset + 1);
    let name = &xml[1..name_end];
    let mut output = Vec::with_capacity(xml.len() + name.len() + 3);
    output.extend_from_slice(&xml[..opening_end - 2]);
    output.push(b'>');
    output.extend_from_slice(&xml[opening_end..]);
    output.extend_from_slice(b"</");
    output.extend_from_slice(name);
    output.push(b'>');
    Ok(output)
}

fn rebind_candidate_child(
    candidate: &[u8],
    original_section: &[u8],
    fragment: &FragmentContext,
) -> Result<Vec<u8>> {
    let candidate_end = opening_tag_end(candidate)?;
    let candidate_prefix = element_prefix(&candidate[..candidate_end]);
    let original_end = opening_tag_end(original_section)?;
    let target_prefix = element_prefix(&original_section[..original_end]);
    let candidate_attributes = lexical_attributes(&candidate[..candidate_end])?;
    let candidate_attribute = candidate_attributes
        .iter()
        .find(|attribute| !is_namespace_attribute(&attribute.prefix, &attribute.local));
    let has_attributes = candidate_attribute.is_some();
    let candidate_attribute_prefix =
        candidate_attribute.map_or(candidate_prefix, |attribute| attribute.prefix.as_slice());
    let mut target_attribute_prefix = Vec::new();
    let mut bind_attribute_namespace = None;
    if has_attributes {
        if !target_prefix.is_empty() {
            target_attribute_prefix.extend_from_slice(target_prefix);
        } else if !candidate_attribute_prefix.is_empty()
            && let Some((prefix, _namespace)) =
                fragment.in_scope_word_attribute_prefix(0, Some(candidate_attribute_prefix))?
        {
            target_attribute_prefix = prefix;
        } else if let Some((prefix, _namespace)) =
            fragment.in_scope_word_attribute_prefix(0, None)?
        {
            target_attribute_prefix = prefix;
        } else {
            let (prefix, namespace) = fragment.fresh_word_attribute_prefix(0)?;
            target_attribute_prefix = prefix;
            bind_attribute_namespace = Some((target_attribute_prefix.clone(), namespace));
        }
    }
    let mut output = rewrite_candidate_qnames(
        candidate,
        candidate_prefix,
        target_prefix,
        candidate_attribute_prefix,
        &target_attribute_prefix,
    )?;
    if let Some((prefix, namespace)) = bind_attribute_namespace {
        output = bind_candidate_attribute_namespace(&output, &prefix, &namespace)?;
    }
    Ok(output)
}

fn rewrite_candidate_qnames(
    candidate: &[u8],
    candidate_element_prefix: &[u8],
    target_element_prefix: &[u8],
    candidate_attribute_prefix: &[u8],
    target_attribute_prefix: &[u8],
) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(candidate.len());
    let mut index = 0usize;
    while index < candidate.len() {
        if candidate[index] != b'<' {
            output.push(candidate[index]);
            index += 1;
            continue;
        }
        let end = index
            .checked_add(opening_tag_end(&candidate[index..])?)
            .ok_or_else(|| Error::InvalidFormat("section candidate tag offset overflow".into()))?;
        let tag = candidate.get(index..end).ok_or_else(|| {
            Error::InvalidFormat("section candidate tag range is outside XML".into())
        })?;
        if tag.starts_with(b"<!--") || tag.starts_with(b"<![CDATA[") || tag.starts_with(b"<?") {
            output.extend_from_slice(tag);
        } else {
            output.extend_from_slice(&rewrite_candidate_tag(
                tag,
                candidate_element_prefix,
                target_element_prefix,
                candidate_attribute_prefix,
                target_attribute_prefix,
            )?);
        }
        index = end;
    }
    Ok(output)
}

fn rewrite_candidate_tag(
    tag: &[u8],
    candidate_element_prefix: &[u8],
    target_element_prefix: &[u8],
    candidate_attribute_prefix: &[u8],
    target_attribute_prefix: &[u8],
) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(tag.len());
    let mut cursor = 0usize;
    if tag.first() != Some(&b'<') {
        return Err(Error::InvalidFormat(
            "section candidate tag does not start with '<'".into(),
        ));
    }
    output.push(b'<');
    cursor += 1;
    if tag.get(cursor) == Some(&b'/') {
        output.push(b'/');
        cursor += 1;
    }
    let name_start = cursor;
    while cursor < tag.len()
        && !tag[cursor].is_ascii_whitespace()
        && !matches!(tag[cursor], b'/' | b'>')
    {
        cursor += 1;
    }
    if name_start == cursor {
        return Err(Error::InvalidFormat(
            "section candidate tag has no qualified name".into(),
        ));
    }
    output.extend_from_slice(&rewrite_element_qname(
        &tag[name_start..cursor],
        candidate_element_prefix,
        target_element_prefix,
    ));
    while cursor < tag.len() {
        if tag[cursor] == b'>' {
            output.push(b'>');
            return Ok(output);
        }
        if tag[cursor].is_ascii_whitespace() || tag[cursor] == b'/' {
            output.push(tag[cursor]);
            cursor += 1;
            continue;
        }
        let attribute_start = cursor;
        while cursor < tag.len()
            && !tag[cursor].is_ascii_whitespace()
            && !matches!(tag[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        if attribute_start == cursor {
            return Err(Error::InvalidFormat(
                "section candidate tag has an invalid attribute name".into(),
            ));
        }
        let attribute_name = &tag[attribute_start..cursor];
        if is_namespace_qname(attribute_name) {
            output.extend_from_slice(attribute_name);
        } else {
            let attribute_prefix = qname_prefix(attribute_name);
            if attribute_prefix.is_empty() || attribute_prefix == candidate_attribute_prefix {
                output.extend_from_slice(&rewrite_attribute_qname(
                    attribute_name,
                    target_attribute_prefix,
                ));
            } else {
                output.extend_from_slice(attribute_name);
            }
        }
        let spacing_start = cursor;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        output.extend_from_slice(&tag[spacing_start..cursor]);
        if tag.get(cursor) != Some(&b'=') {
            continue;
        }
        output.push(b'=');
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        output.extend_from_slice(&tag[value_start..cursor]);
        let Some(&quote) = tag.get(cursor) else {
            return Err(Error::InvalidFormat(
                "section candidate attribute value is missing".into(),
            ));
        };
        if quote != b'\'' && quote != b'"' {
            return Err(Error::InvalidFormat(
                "section candidate attribute value is not quoted".into(),
            ));
        }
        let quoted_start = cursor;
        cursor += 1;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        if cursor >= tag.len() {
            return Err(Error::InvalidFormat(
                "section candidate attribute value is unterminated".into(),
            ));
        }
        cursor += 1;
        output.extend_from_slice(&tag[quoted_start..cursor]);
    }
    Err(Error::InvalidFormat(
        "section candidate tag has no closing '>'".into(),
    ))
}

fn rewrite_element_qname(name: &[u8], candidate_prefix: &[u8], target_prefix: &[u8]) -> Vec<u8> {
    let prefix = qname_prefix(name);
    if prefix != candidate_prefix {
        return name.to_vec();
    }
    let local = qname_local(name);
    let mut output = Vec::with_capacity(target_prefix.len() + local.len() + 1);
    if !target_prefix.is_empty() {
        output.extend_from_slice(target_prefix);
        output.push(b':');
    }
    output.extend_from_slice(local);
    output
}

fn rewrite_attribute_qname(name: &[u8], target_prefix: &[u8]) -> Vec<u8> {
    let local = qname_local(name);
    let mut output = Vec::with_capacity(target_prefix.len() + local.len() + 1);
    if !target_prefix.is_empty() {
        output.extend_from_slice(target_prefix);
        output.push(b':');
    }
    output.extend_from_slice(local);
    output
}

fn qname_prefix(name: &[u8]) -> &[u8] {
    name.iter()
        .position(|byte| *byte == b':')
        .map_or(b"".as_slice(), |index| &name[..index])
}

fn qname_local(name: &[u8]) -> &[u8] {
    name.iter()
        .position(|byte| *byte == b':')
        .map_or(name, |index| &name[index + 1..])
}

fn is_namespace_qname(name: &[u8]) -> bool {
    name == b"xmlns" || name.strip_prefix(b"xmlns:").is_some()
}

fn is_namespace_attribute(prefix: &[u8], local: &[u8]) -> bool {
    prefix == b"xmlns" || (prefix.is_empty() && local == b"xmlns")
}

fn bind_candidate_attribute_namespace(
    candidate: &[u8],
    prefix: &[u8],
    namespace: &[u8],
) -> Result<Vec<u8>> {
    if prefix.is_empty() {
        return Err(Error::UnsafeEdit {
            format: "DOCX",
            operation: "edit_section_layout",
            reason: "the inserted section property has no safe Word attribute prefix",
        });
    }
    let opening_end = opening_tag_end(candidate)?;
    let attributes = lexical_attributes(&candidate[..opening_end])?;
    if let Some(existing) = attributes
        .iter()
        .find(|attribute| attribute.prefix == b"xmlns" && attribute.local.as_slice() == prefix)
    {
        if &candidate[existing.value_start..existing.value_end] == namespace {
            return Ok(candidate.to_vec());
        }
        return Err(Error::UnsafeEdit {
            format: "DOCX",
            operation: "edit_section_layout",
            reason: "the inserted section property has a conflicting Word attribute namespace",
        });
    }
    let insertion = if candidate[..opening_end].ends_with(b"/>") {
        opening_end - 2
    } else {
        opening_end - 1
    };
    let mut declaration = Vec::new();
    declaration.extend_from_slice(b" xmlns:");
    declaration.extend_from_slice(prefix);
    declaration.extend_from_slice(b"=\"");
    declaration.extend_from_slice(namespace);
    declaration.extend_from_slice(b"\"");
    replace_range(candidate, insertion, insertion, &declaration)
}

fn is_word_namespace_bytes(namespace: &[u8]) -> bool {
    matches!(
        namespace,
        b"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            | b"http://purl.oclc.org/ooxml/wordprocessingml/main"
    )
}

fn opening_tag_end(xml: &[u8]) -> Result<usize> {
    let mut quote = None;
    for (index, byte) in xml.iter().enumerate() {
        if let Some(delimiter) = quote {
            if *byte == delimiter {
                quote = None;
            }
        } else if matches!(*byte, b'\'' | b'"') {
            quote = Some(*byte);
        } else if *byte == b'>' {
            return Ok(index + 1);
        }
    }
    Err(Error::InvalidFormat(
        "section opening element is incomplete".into(),
    ))
}

fn element_prefix(opening: &[u8]) -> &[u8] {
    let mut end = 1usize;
    while end < opening.len()
        && !opening[end].is_ascii_whitespace()
        && !matches!(opening[end], b'>' | b'/')
    {
        end += 1;
    }
    let name = &opening[1..end];
    name.iter()
        .position(|byte| *byte == b':')
        .map_or(b"".as_slice(), |index| &name[..index])
}

fn replace_range(source: &[u8], start: usize, end: usize, replacement: &[u8]) -> Result<Vec<u8>> {
    let prefix = source
        .get(..start)
        .ok_or_else(|| Error::InvalidFormat("section source range starts outside XML".into()))?;
    let suffix = source
        .get(end..)
        .ok_or_else(|| Error::InvalidFormat("section source range ends outside XML".into()))?;
    let capacity = prefix
        .len()
        .checked_add(replacement.len())
        .and_then(|size| size.checked_add(suffix.len()))
        .ok_or_else(|| Error::InvalidFormat("section layout output size overflow".into()))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "section layout document XML",
            source,
        })?;
    output.extend_from_slice(prefix);
    output.extend_from_slice(replacement);
    output.extend_from_slice(suffix);
    Ok(output)
}

fn ensure_same_topology(before: &Inventory, after: &Inventory) -> Result<()> {
    if !before.topology_matches(after) {
        return Err(Error::UnsafeEdit {
            format: "DOCX",
            operation: "edit_section_layout",
            reason: "section layout edit changed main-story topology",
        });
    }
    Ok(())
}

fn ensure_semantic_readback(
    inventory: &Inventory,
    selector: Selector,
    expected: &LocalState,
) -> Result<()> {
    let section = inventory.section(selector).ok_or(Error::UnsafeEdit {
        format: "DOCX",
        operation: "edit_section_layout",
        reason: "section layout edit lost its selected main-story section",
    })?;
    if section.page_size() != expected.page_size
        || section.margins() != expected.margins
        || section.start() != expected.start
        || section.columns() != expected.columns
    {
        return Err(Error::UnsafeEdit {
            format: "DOCX",
            operation: "edit_section_layout",
            reason: "section layout semantic readback differs from the projected edit",
        });
    }
    Ok(())
}
