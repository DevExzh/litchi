//! Exact source-backed slide-order edits.
//!
//! The order capability changes only the ordered `p:sldId` entries in the
//! existing presentation part. It retains every entry as an opaque span and
//! refuses presentations whose other known positional owners would become
//! stale after a move.

use std::io::Write;
use std::sync::Arc;

use litchi_core::{ExecutionContext, SourceVersion};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{PackURI, PartData, Relationships, SourceLineage, TargetMode};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::source::SourceBackedPresentationEditor;
use crate::{Error, Result, SlideOrderRefusal};

const MAX_XML_DEPTH: usize = 256;
const MAX_XML_NODES: usize = 1_000_000;

const PRESENTATIONML_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PRESENTATIONML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const STRICT_MCE_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/markup-compatibility/2006";
const CANONICAL_PRESENTATION_MEMBER: &str = "ppt/presentation.xml";

const STRICT_PRES_PROPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/presProps";
const STRICT_VIEW_PROPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/viewProps";

#[derive(Clone, PartialEq, Eq)]
struct SlideOrderBinding {
    id: u32,
    relationship_id: String,
}

#[derive(Clone, PartialEq, Eq)]
struct RootRelationshipBinding {
    id: String,
    kind: String,
    target: String,
    mode: TargetMode,
}

/// A retained presentation XML payload for an order snapshot.
///
/// Original bytes remain attached to the OPC `PartData` reservation. A
/// changed candidate is edit-owned and is retained through shared immutable
/// bytes until publication.
#[derive(Clone)]
enum OrderPayload {
    Original(PartData),
    Edited(Arc<Vec<u8>>),
}

impl OrderPayload {
    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Original(data) => data.as_bytes(),
            Self::Edited(data) => data.as_slice(),
        }
    }

    fn edited_bytes(&self) -> Option<Arc<Vec<u8>>> {
        match self {
            Self::Original(_) => None,
            Self::Edited(data) => Some(Arc::clone(data)),
        }
    }
}

/// Exact source-backed presentation order captured before one edit.
///
/// The snapshot is process-local. Application requires the same source
/// lineage and version, presentation URI and bytes, root relationship
/// bindings, and ordered `(p:sldId@id, p:sldId@r:id)` bindings.
#[derive(Clone)]
pub struct SourceBackedSlideOrderSnapshot {
    presentation_uri: PackURI,
    xml: OrderPayload,
    bindings: Box<[SlideOrderBinding]>,
    relationships: Box<[RootRelationshipBinding]>,
    source_version: SourceVersion,
    lineage: SourceLineage,
    context: Option<ExecutionContext>,
}

/// An isolated source-backed edit that can stage exactly one positional move.
pub struct SourceBackedSlideOrderEdit {
    source: SourceBackedSlideOrderSnapshot,
    working: OrderPayload,
    bindings: Box<[SlideOrderBinding]>,
    operation_used: bool,
}

/// A reversible exact-source presentation-order patch.
#[derive(Clone)]
pub struct SourceBackedSlideOrderPatch {
    before: SourceBackedSlideOrderSnapshot,
    after: SourceBackedSlideOrderSnapshot,
}

/// A checked source-backed presentation-order edit ready for publication.
pub struct SourceBackedSlideOrderCommit {
    snapshot: SourceBackedSlideOrderSnapshot,
    patch: SourceBackedSlideOrderPatch,
}

impl SourceBackedPresentationEditor {
    /// Begin an isolated exact-source presentation-order edit.
    ///
    /// The returned edit performs no slide or media payload reads. Call
    /// [`SourceBackedSlideOrderEdit::move_slide`] at most once. The same
    /// position is an explicit exact no-op.
    pub fn edit_slide_order(&self) -> Result<SourceBackedSlideOrderEdit> {
        let source = capture(self)?;
        Ok(SourceBackedSlideOrderEdit {
            working: source.xml.clone(),
            bindings: source.bindings.clone(),
            source,
            operation_used: false,
        })
    }

    /// Begin an order edit and stage one positional move.
    ///
    /// This convenience method is equivalent to [`Self::edit_slide_order`]
    /// followed by [`SourceBackedSlideOrderEdit::move_slide`].
    pub fn edit_slide_order_move(
        &self,
        from: usize,
        to: usize,
    ) -> Result<SourceBackedSlideOrderEdit> {
        let mut edit = self.edit_slide_order()?;
        edit.move_slide(from, to)?;
        Ok(edit)
    }

    /// Publish an exact-source-checked slide-order commit to a sequential
    /// stream.
    pub fn publish_slide_order_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &SourceBackedSlideOrderCommit,
    ) -> Result<SourceBackedSlideOrderSnapshot> {
        self.publish_slide_order_patch_to_stream(writer, commit.patch())
    }

    /// Publish an exact-source-checked slide-order patch to a sequential
    /// stream.
    ///
    /// An effective move overlays only the existing presentation XML part.
    /// An exact no-op selects the empty overlay path, preserving every source
    /// ZIP member byte-for-byte, including signature infrastructure.
    pub fn publish_slide_order_patch_to_stream<W: Write>(
        self,
        writer: W,
        patch: &SourceBackedSlideOrderPatch,
    ) -> Result<SourceBackedSlideOrderSnapshot> {
        self.package.check_execution()?;
        let current = capture(&self)?;
        let target = patch.apply(&current)?;
        self.package.check_execution()?;
        if patch.is_changed() {
            validate_for_rewrite(&current)?;
            validate_effective_publication_source(&self.package, &current.presentation_uri)?;
            let replacement = target.xml.edited_bytes().ok_or_else(|| {
                Error::Invalid("source-backed slide-order target has no replacement".into())
            })?;
            self.package.write_part_overlay_shared_to_stream(
                writer,
                &current.presentation_uri,
                replacement,
            )?;
        } else {
            self.package
                .write_part_overlays_shared_to_stream(writer, Vec::new())?;
        }
        Ok(target)
    }
}

impl SourceBackedSlideOrderSnapshot {
    /// Number of ordered slide entries.
    #[must_use]
    pub fn slide_count(&self) -> usize {
        self.bindings.len()
    }

    /// Whether this snapshot is exactly the same source snapshot as `other`.
    #[must_use]
    pub fn is_exactly_same_as(&self, other: &Self) -> bool {
        self.same_source(other)
    }

    fn check_execution(&self) -> Result<()> {
        check_execution_context(self.context.as_ref())
    }

    fn same_source(&self, other: &Self) -> bool {
        self.presentation_uri == other.presentation_uri
            && self.xml.as_bytes() == other.xml.as_bytes()
            && self.bindings == other.bindings
            && self.relationships == other.relationships
            && self.source_version == other.source_version
            && self.lineage == other.lineage
    }
}

impl SourceBackedSlideOrderEdit {
    /// Exact immutable presentation order captured at edit start.
    #[must_use]
    pub const fn source(&self) -> &SourceBackedSlideOrderSnapshot {
        &self.source
    }

    /// Whether this edit has a byte-changing move staged.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.source.xml.as_bytes() != self.working.as_bytes()
    }

    /// Alias for [`Self::is_changed`].
    #[must_use]
    pub fn changed(&self) -> bool {
        self.is_changed()
    }

    /// Stage exactly one positional move in the ordered slide list.
    ///
    /// `from == to` is a successful, explicit exact no-op and consumes this
    /// edit's one operation. A second move is refused before any mutation.
    pub fn move_slide(&mut self, from: usize, to: usize) -> Result<bool> {
        self.source.check_execution()?;
        if self.operation_used {
            return Err(refusal(
                SlideOrderRefusal::MultipleMoves,
                "source-backed slide-order edits support one positional move",
            ));
        }
        let length = self.source.bindings.len();
        if from >= length {
            return Err(Error::SlideIndexOutOfBounds {
                index: from,
                len: length,
            });
        }
        if to >= length {
            return Err(Error::SlideIndexOutOfBounds {
                index: to,
                len: length,
            });
        }
        if from == to {
            self.source.check_execution()?;
            self.operation_used = true;
            return Ok(false);
        }

        validate_for_rewrite(&self.source)?;
        let mut ordered = clone_bindings(&self.source.bindings)?;
        let binding = ordered.remove(from);
        ordered.insert(to, binding);
        let current_refs = binding_refs(&self.source.bindings)?;
        let ordered_refs = binding_refs(&ordered)?;
        let xml = crate::opened::reorder_slide_bindings(
            self.source.xml.as_bytes(),
            &current_refs,
            &ordered_refs,
            self.source.context.as_ref(),
        )?;
        let candidate_working = OrderPayload::Edited(Arc::new(xml));
        let candidate_bindings = ordered.into_boxed_slice();
        self.source.check_execution()?;
        self.working = candidate_working;
        self.bindings = candidate_bindings;
        self.operation_used = true;
        Ok(true)
    }

    /// Alias for [`Self::move_slide`].
    pub fn move_position(&mut self, from: usize, to: usize) -> Result<bool> {
        self.move_slide(from, to)
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    pub fn commit(self) -> Result<SourceBackedSlideOrderCommit> {
        self.source.check_execution()?;
        if !self.operation_used {
            return Err(refusal(
                SlideOrderRefusal::NoMove,
                "stage one positional move before committing",
            ));
        }
        Ok(self.into_commit())
    }

    fn into_commit(self) -> SourceBackedSlideOrderCommit {
        let snapshot = SourceBackedSlideOrderSnapshot {
            presentation_uri: self.source.presentation_uri.clone(),
            xml: self.working,
            bindings: self.bindings,
            relationships: self.source.relationships.clone(),
            source_version: self.source.source_version.clone(),
            lineage: self.source.lineage.clone(),
            context: self.source.context.clone(),
        };
        let patch = SourceBackedSlideOrderPatch {
            before: self.source,
            after: snapshot.clone(),
        };
        SourceBackedSlideOrderCommit { snapshot, patch }
    }
}

impl SourceBackedSlideOrderPatch {
    /// Exact immutable source required by this patch.
    #[must_use]
    pub const fn source(&self) -> &SourceBackedSlideOrderSnapshot {
        &self.before
    }

    /// Exact immutable target produced by this patch.
    #[must_use]
    pub const fn target(&self) -> &SourceBackedSlideOrderSnapshot {
        &self.after
    }

    /// Whether this patch changes the presentation XML bytes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.before.same_source(&self.after)
    }

    /// Alias for [`Self::is_changed`].
    #[must_use]
    pub fn changed(&self) -> bool {
        self.is_changed()
    }

    /// Whether this patch is an exact source no-op.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.is_changed()
    }

    /// Return the exact source-bound inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply only to the exact source captured by this patch.
    pub fn apply(
        &self,
        source: &SourceBackedSlideOrderSnapshot,
    ) -> Result<SourceBackedSlideOrderSnapshot> {
        self.before.check_execution()?;
        source.check_execution()?;
        if !source.same_source(&self.before) {
            return Err(Error::StaleSource);
        }
        let target = if self.is_changed() {
            self.after.clone()
        } else {
            source.clone()
        };
        target.check_execution()?;
        Ok(target)
    }
}

impl SourceBackedSlideOrderCommit {
    /// Candidate snapshot after this edit.
    #[must_use]
    pub const fn snapshot(&self) -> &SourceBackedSlideOrderSnapshot {
        &self.snapshot
    }

    /// Exact reversible patch for this edit.
    #[must_use]
    pub const fn patch(&self) -> &SourceBackedSlideOrderPatch {
        &self.patch
    }

    /// Whether the ordered slide entries change.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.patch.is_changed()
    }

    /// Alias for [`Self::is_changed`].
    #[must_use]
    pub fn changed(&self) -> bool {
        self.is_changed()
    }

    /// Exact reversible inverse patch for this edit.
    #[must_use]
    pub fn inverse(&self) -> SourceBackedSlideOrderPatch {
        self.patch.inverse()
    }
}

fn capture(editor: &SourceBackedPresentationEditor) -> Result<SourceBackedSlideOrderSnapshot> {
    editor.package.check_execution()?;
    let source_version = editor.package.source_version()?;
    let view = editor.package.main_document_part()?;
    let raw = editor._presentation.source_data().clone();
    editor.package.check_execution()?;
    if view.partname() != editor._presentation.source_partname() {
        return Err(Error::StaleSource);
    }

    let bindings = slide_order_bindings(editor)?;
    let relationships = root_relationship_bindings(view.rels())?;
    let pinned_relationships =
        root_relationship_bindings(editor._presentation.source_relationships())?;
    if relationships != pinned_relationships {
        return Err(Error::StaleSource);
    }
    editor.package.source_version()?;
    let lineage = editor.package.source_lineage();
    let context = editor.package.execution_context();
    Ok(SourceBackedSlideOrderSnapshot {
        presentation_uri: view.partname().clone(),
        xml: OrderPayload::Original(raw),
        bindings,
        relationships,
        source_version,
        lineage,
        context,
    })
}

fn validate_for_rewrite(snapshot: &SourceBackedSlideOrderSnapshot) -> Result<()> {
    snapshot.check_execution()?;
    reject_root_features(snapshot.xml.as_bytes(), snapshot.context.as_ref())?;
    let binding_refs = binding_refs(&snapshot.bindings)?;
    crate::opened::validate_slide_bindings(
        snapshot.xml.as_bytes(),
        &binding_refs,
        snapshot.context.as_ref(),
    )
    .map_err(|error| match error {
        Error::Invalid(detail) => refusal(SlideOrderRefusal::AmbiguousBindings, detail),
        other => other,
    })?;
    snapshot.check_execution()?;
    Ok(())
}

fn validate_effective_publication_source(
    package: &litchi_opc::SourceBackedPackage,
    presentation_uri: &PackURI,
) -> Result<()> {
    package.validate_topology_source_boundary()?;
    if presentation_uri.as_str() != "/ppt/presentation.xml" {
        return Err(Error::Opc(
            litchi_opc::OpcError::SourceBackedOverlayUnavailable {
                reason: "slide-order publication requires the canonical /ppt/presentation.xml part"
                    .to_owned(),
            },
        ));
    }
    let mut exact_member = false;
    for name in package.physical_member_names() {
        if name.eq_ignore_ascii_case(CANONICAL_PRESENTATION_MEMBER) {
            if name != CANONICAL_PRESENTATION_MEMBER || exact_member {
                return Err(Error::Opc(
                    litchi_opc::OpcError::SourceBackedOverlayUnavailable {
                        reason: "slide-order publication refuses case-equivalent presentation.xml physical aliases"
                            .to_owned(),
                    },
                ));
            }
            exact_member = true;
        }
    }
    if !exact_member {
        return Err(Error::Opc(
            litchi_opc::OpcError::SourceBackedOverlayUnavailable {
                reason:
                    "slide-order publication requires an exact ppt/presentation.xml physical member"
                        .to_owned(),
            },
        ));
    }
    if package.has_encrypted_entries() {
        return Err(Error::Opc(
            litchi_opc::OpcError::SourceBackedOverlayUnavailable {
                reason: "slide-order publication refuses encrypted ZIP entries".to_owned(),
            },
        ));
    }
    validate_positional_metadata(package)?;
    Ok(())
}

#[derive(Clone, Copy)]
enum MetadataOwner {
    ViewProperties,
    PresentationProperties,
}

fn validate_positional_metadata(package: &litchi_opc::SourceBackedPackage) -> Result<()> {
    let context = package.execution_context();
    let presentation = package.main_document_part()?;
    for relationship in presentation.rels().iter() {
        check_execution_context(context.as_ref())?;
        let owner = match relationship.reltype() {
            rt::VIEW_PROPS | STRICT_VIEW_PROPS => MetadataOwner::ViewProperties,
            rt::PRES_PROPS | STRICT_PRES_PROPS => MetadataOwner::PresentationProperties,
            _ => continue,
        };
        if relationship.is_external() {
            return Err(metadata_refusal(
                owner,
                "positional metadata relationship is external",
            ));
        }
        let partname = relationship.target_partname()?;
        let data = package.part(&partname)?.data()?;
        inspect_metadata_xml(data.as_bytes(), owner, context.as_ref())?;
        check_execution_context(context.as_ref())?;
        match owner {
            MetadataOwner::ViewProperties => {
                let properties = crate::view_properties::ViewProperties::parse(data.as_bytes())
                    .map_err(|error| {
                        metadata_refusal(owner, format!("view properties parse failed: {error}"))
                    })?;
                if properties
                    .outline
                    .as_ref()
                    .is_some_and(|outline| !outline.slides.is_empty())
                {
                    return Err(metadata_refusal(
                        owner,
                        "view properties contain outline slide-position metadata",
                    ));
                }
            },
            MetadataOwner::PresentationProperties => {
                let properties = crate::presentation_properties::Properties::parse(data.as_bytes())
                    .map_err(|error| {
                        metadata_refusal(
                            owner,
                            format!("presentation properties parse failed: {error}"),
                        )
                    })?;
                let html_has_range = properties.html_publish.as_ref().is_some_and(|html| {
                    matches!(
                        &html.slides,
                        crate::presentation_properties::SlideSelection::Range { .. }
                    )
                });
                let show_has_range = properties.show.as_ref().is_some_and(|show| {
                    show.slides.as_ref().is_some_and(|selection| {
                        matches!(
                            selection,
                            crate::presentation_properties::SlideSelection::Range { .. }
                        )
                    })
                });
                if html_has_range || show_has_range {
                    return Err(metadata_refusal(
                        owner,
                        "presentation properties contain a numeric slide-position range",
                    ));
                }
            },
        }
        check_execution_context(context.as_ref())?;
    }
    check_execution_context(context.as_ref())?;
    Ok(())
}

fn inspect_metadata_xml(
    xml: &[u8],
    owner: MetadataOwner,
    context: Option<&ExecutionContext>,
) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut nodes = 0usize;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                bump_xml(&mut nodes)?;
                if nodes % 256 == 0 {
                    check_execution_context(context)?;
                }
                if is_mce_namespace(&namespace) || has_mce_attribute(&element)? {
                    return Err(metadata_refusal(
                        owner,
                        "positional metadata contains markup-compatibility content",
                    ));
                }
                let local_name = element.name().local_name();
                if local_name.as_ref() == b"extLst" || local_name.as_ref() == b"ext" {
                    return Err(metadata_refusal(
                        owner,
                        "positional metadata contains extension content",
                    ));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(metadata_refusal(
                    owner,
                    "positional metadata contains a DTD or processing instruction",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    check_execution_context(context)?;
    Ok(())
}

fn metadata_refusal(owner: MetadataOwner, detail: impl Into<String>) -> Error {
    refusal(
        match owner {
            MetadataOwner::ViewProperties => SlideOrderRefusal::ViewProperties,
            MetadataOwner::PresentationProperties => SlideOrderRefusal::PresentationProperties,
        },
        detail,
    )
}

fn slide_order_bindings(
    editor: &SourceBackedPresentationEditor,
) -> Result<Box<[SlideOrderBinding]>> {
    let mut bindings = Vec::new();
    bindings
        .try_reserve_exact(editor.slides.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed slide-order bindings",
            source,
        })?;
    for slide in &editor.slides {
        bindings.push(SlideOrderBinding {
            id: slide.slide_id,
            relationship_id: slide.binding.slide_reference_id.clone(),
        });
    }
    Ok(bindings.into_boxed_slice())
}

fn clone_bindings(bindings: &[SlideOrderBinding]) -> Result<Vec<SlideOrderBinding>> {
    let mut cloned = Vec::new();
    cloned
        .try_reserve_exact(bindings.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed ordered slide bindings",
            source,
        })?;
    cloned.extend(bindings.iter().cloned());
    Ok(cloned)
}

fn binding_refs<'a>(bindings: &'a [SlideOrderBinding]) -> Result<Vec<(u32, &'a str)>> {
    let mut refs = Vec::new();
    refs.try_reserve_exact(bindings.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed slide-order binding references",
            source,
        })?;
    for binding in bindings {
        refs.push((binding.id, binding.relationship_id.as_str()));
    }
    Ok(refs)
}

fn root_relationship_bindings(
    relationships: &Relationships,
) -> Result<Box<[RootRelationshipBinding]>> {
    let mut bindings = Vec::new();
    bindings
        .try_reserve_exact(relationships.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed presentation relationships",
            source,
        })?;
    for relationship in relationships.iter() {
        bindings.push(RootRelationshipBinding {
            id: relationship.r_id().to_owned(),
            kind: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            mode: relationship.target_mode(),
        });
    }
    bindings.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    Ok(bindings.into_boxed_slice())
}

fn reject_root_features(xml: &[u8], context: Option<&ExecutionContext>) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut roots = 0usize;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                bump_xml(&mut nodes)?;
                if nodes % 256 == 0 {
                    check_execution_context(context)?;
                }
                if depth >= MAX_XML_DEPTH {
                    return Err(Error::Limit {
                        resource: "source-backed presentation XML depth",
                        limit: MAX_XML_DEPTH,
                    });
                }
                reject_mce_element(&namespace, &element)?;
                if depth == 0 {
                    roots = roots
                        .checked_add(1)
                        .ok_or_else(|| Error::Invalid("presentation XML root overflow".into()))?;
                    if !is_presentation_name(&namespace, element.name(), b"presentation") {
                        return Err(Error::Invalid(
                            "source-backed presentation XML has an unexpected root".into(),
                        ));
                    }
                    validate_root_attributes(&element)?;
                } else if depth == 1 {
                    reject_root_child(&namespace, element.name().local_name().as_ref())?;
                }
                reject_positional_element(element.name().local_name().as_ref())?;
                depth += 1;
            },
            Event::Empty(element) => {
                bump_xml(&mut nodes)?;
                if nodes % 256 == 0 {
                    check_execution_context(context)?;
                }
                reject_mce_element(&namespace, &element)?;
                if depth == 0 {
                    roots = roots
                        .checked_add(1)
                        .ok_or_else(|| Error::Invalid("presentation XML root overflow".into()))?;
                    if !is_presentation_name(&namespace, element.name(), b"presentation") {
                        return Err(Error::Invalid(
                            "source-backed presentation XML has an unexpected root".into(),
                        ));
                    }
                    validate_root_attributes(&element)?;
                } else if depth == 1 {
                    reject_root_child(&namespace, element.name().local_name().as_ref())?;
                }
                reject_positional_element(element.name().local_name().as_ref())?;
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Invalid("source-backed presentation XML depth underflow".into())
                })?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::Invalid(
                    "source-backed presentation XML contains a DTD or processing instruction"
                        .into(),
                ));
            },
            Event::Text(value) if depth == 0 => {
                let text = value
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                if !text.trim().is_empty() {
                    return Err(Error::Invalid(
                        "source-backed presentation XML contains semantic text outside its root"
                            .into(),
                    ));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(Error::Invalid(
                    "source-backed presentation XML contains semantic content outside its root"
                        .into(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    check_execution_context(context)?;
    if depth != 0 || roots != 1 {
        return Err(Error::Invalid(
            "source-backed presentation XML must contain one closed root".into(),
        ));
    }
    Ok(())
}

fn reject_mce_element(
    namespace: &ResolveResult<'_>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<()> {
    if is_mce_namespace(namespace) {
        return Err(refusal(
            SlideOrderRefusal::MarkupCompatibility,
            "presentation XML contains markup-compatibility content",
        ));
    }
    if has_mce_attribute(element)? {
        return Err(refusal(
            SlideOrderRefusal::MarkupCompatibility,
            "presentation XML contains markup-compatibility attributes",
        ));
    }
    Ok(())
}

fn has_mce_attribute(element: &quick_xml::events::BytesStart<'_>) -> Result<bool> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if matches!(
            attribute.key.local_name().as_ref(),
            b"Ignorable"
                | b"MustUnderstand"
                | b"ProcessContent"
                | b"PreserveAttributes"
                | b"PreserveElements"
        ) {
            return Ok(true);
        }
    }
    Ok(false)
}

fn validate_root_attributes(element: &quick_xml::events::BytesStart<'_>) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        if key.starts_with(b"xml:") || (!key.contains(&b':') && known_root_attribute(key)) {
            continue;
        }
        return Err(refusal(
            SlideOrderRefusal::PositionalOwner,
            "presentation root contains an unmodeled attribute",
        ));
    }
    Ok(())
}

fn known_root_attribute(attribute: &[u8]) -> bool {
    matches!(
        attribute,
        b"saveSubsetFonts"
            | b"autoCompressPictures"
            | b"serverZoom"
            | b"bookmarkIdSeed"
            | b"firstSlideNum"
            | b"showSpecialPlsOnTitleSld"
            | b"rtl"
            | b"removePersonalInfoOnSave"
            | b"compatMode"
            | b"strictFirstAndLastChars"
            | b"conformance"
            | b"embedTrueTypeFonts"
    )
}

fn reject_root_child(namespace: &ResolveResult<'_>, local_name: &[u8]) -> Result<()> {
    if !is_presentation_namespace(namespace) {
        return Err(refusal(
            SlideOrderRefusal::PositionalOwner,
            "presentation root contains a foreign direct child",
        ));
    }
    match local_name {
        b"modifyVerifier" => Err(refusal(
            SlideOrderRefusal::ModifyVerifier,
            "presentation modify verifier is a positional semantic owner",
        )),
        b"extLst" | b"photoAlbum" => Err(refusal(
            SlideOrderRefusal::PositionalOwner,
            "presentation root extension or photo-album metadata may own slide positions",
        )),
        b"sldMasterIdLst"
        | b"notesMasterIdLst"
        | b"handoutMasterIdLst"
        | b"sldIdLst"
        | b"sldSz"
        | b"notesSz"
        | b"smartTags"
        | b"embeddedFontLst"
        | b"custShowLst"
        | b"custDataLst"
        | b"kinsoku"
        | b"defaultTextStyle" => Ok(()),
        _ => Err(refusal(
            SlideOrderRefusal::PositionalOwner,
            "presentation root contains an unmodeled direct child",
        )),
    }
}

fn reject_positional_element(local_name: &[u8]) -> Result<()> {
    if local_name == b"sectionLst" || local_name == b"section" {
        return Err(refusal(
            SlideOrderRefusal::Sections,
            "presentation section metadata owns slide positions",
        ));
    }
    Ok(())
}

fn is_mce_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == MCE_NAMESPACE || *value == STRICT_MCE_NAMESPACE
    )
}

fn check_execution_context(context: Option<&ExecutionContext>) -> Result<()> {
    if let Some(context) = context {
        context.check().map_err(|error| {
            Error::Opc(match error {
                litchi_core::ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                error => litchi_opc::OpcError::Execution(error),
            })
        })?;
    }
    Ok(())
}

fn is_presentation_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == PRESENTATIONML_NAMESPACE || *value == STRICT_PRESENTATIONML_NAMESPACE
    )
}

fn is_presentation_name(
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    local: &[u8],
) -> bool {
    name.local_name().as_ref() == local && is_presentation_namespace(namespace)
}

fn bump_xml(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| Error::Invalid("source-backed presentation XML node overflow".into()))?;
    if *nodes > MAX_XML_NODES {
        return Err(Error::Limit {
            resource: "source-backed presentation XML nodes",
            limit: MAX_XML_NODES,
        });
    }
    Ok(())
}

fn refusal(kind: SlideOrderRefusal, detail: impl Into<String>) -> Error {
    Error::SlideOrderPlan {
        kind,
        detail: detail.into(),
    }
}
