//! Exact source-backed slide-order edits.
//!
//! The order capability changes only the ordered `p:sldId` entries in the
//! existing presentation part. It retains every entry as an opaque span and
//! refuses presentations whose other known positional owners would become
//! stale after a move.

#[cfg(test)]
use std::io::Write;
use std::sync::Arc;

use litchi_core::{ExecutionContext, ExecutionError, Reservation, Resource};
#[cfg(test)]
use litchi_core::{Position, SourceVersion};
use litchi_opc::PackURI;
use litchi_opc::constants::relationship_type as rt;
#[cfg(test)]
use litchi_opc::{PartData, Relationships, SourceLineage, TargetMode};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::source::SourceBackedPresentationEditor;
use crate::{Error, Result, SlideOrderRefusal};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

#[cfg(test)]
macro_rules! assert_refused_kind {
    ($operation:expr, $expected:expr) => {{
        match $operation {
            Err(Error::Refused(refusal)) => {
                assert_eq!(refusal.kind(), $expected);
            },
            Err(error) => panic!("unexpected error: {error:?}"),
            Ok(_) => panic!("operation unexpectedly succeeded"),
        }
    }};
}

const MAX_XML_DEPTH: usize = 256;
const MAX_XML_NODES: usize = 1_000_000;

const PRESENTATIONML_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PRESENTATIONML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const POWERPOINT_2010_NAMESPACE: &[u8] =
    b"http://schemas.microsoft.com/office/powerpoint/2010/main";
const DRAWINGML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const CANONICAL_PRESENTATION_MEMBER: &str = "ppt/presentation.xml";

const STRICT_PRES_PROPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/presProps";
const STRICT_VIEW_PROPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/viewProps";

/// The source identity of one slide in the retained presentation catalog.
///
/// A position is deliberately absent from this value.  Positions are logical
/// and can change when an order edit is staged; the part URI, slide id, and
/// presentation relationship id continue to identify the same physical
/// source slide across that edit.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub(super) struct StableSlideKey {
    pub(super) part_uri: PackURI,
    pub(super) slide_id: u32,
    pub(super) relationship_id: String,
}

/// Construct a stable source identity for a catalog entry.
pub(super) fn stable_slide_key(
    part_uri: PackURI,
    slide_id: u32,
    relationship_id: impl Into<String>,
) -> StableSlideKey {
    StableSlideKey {
        part_uri,
        slide_id,
        relationship_id: relationship_id.into(),
    }
}

#[derive(Clone, PartialEq, Eq)]
#[cfg(test)]
struct SlideOrderBinding {
    id: u32,
    relationship_id: String,
}

#[derive(Clone, PartialEq, Eq)]
#[cfg(test)]
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
#[cfg(test)]
enum OrderPayload {
    Original(PartData),
    Edited(Arc<Vec<u8>>),
    ReservedEdit {
        bytes: Arc<Vec<u8>>,
        _reservation: Arc<Reservation>,
    },
}

#[cfg(test)]
impl OrderPayload {
    #[cfg(test)]
    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Original(data) => data.as_bytes(),
            Self::Edited(data) | Self::ReservedEdit { bytes: data, .. } => data.as_slice(),
        }
    }

    #[cfg(test)]
    fn edited_bytes(&self) -> Option<Arc<Vec<u8>>> {
        match self {
            Self::Original(_) => None,
            Self::Edited(data) | Self::ReservedEdit { bytes: data, .. } => Some(Arc::clone(data)),
        }
    }
}

/// Exact source-backed presentation order captured before one edit.
///
/// The snapshot is process-local. Application requires the same source
/// lineage and version, presentation URI and bytes, root relationship
/// bindings, and ordered `(p:sldId@id, p:sldId@r:id)` bindings.
#[derive(Clone)]
#[cfg(test)]
pub(crate) struct SourceBackedSlideOrderSnapshot {
    presentation_uri: PackURI,
    xml: OrderPayload,
    bindings: Box<[SlideOrderBinding]>,
    relationships: Box<[RootRelationshipBinding]>,
    source_version: SourceVersion,
    lineage: SourceLineage,
    context: Option<ExecutionContext>,
}

/// An isolated source-backed edit that can stage a sequence of positional
/// moves and swaps.
#[cfg(test)]
pub(crate) struct SourceBackedSlideOrderEdit {
    source: SourceBackedSlideOrderSnapshot,
    working: OrderPayload,
    bindings: Box<[SlideOrderBinding]>,
}

/// A reversible exact-source presentation-order patch.
#[derive(Clone)]
#[cfg(test)]
pub(crate) struct SourceBackedSlideOrderPatch {
    before: SourceBackedSlideOrderSnapshot,
    after: SourceBackedSlideOrderSnapshot,
}

/// A checked source-backed presentation-order edit ready for publication.
#[cfg(test)]
pub(crate) struct SourceBackedSlideOrderCommit {
    snapshot: SourceBackedSlideOrderSnapshot,
    patch: SourceBackedSlideOrderPatch,
}

impl SourceBackedPresentationEditor {
    /// Begin an isolated exact-source presentation-order edit.
    ///
    /// The returned edit performs no slide or media payload reads. Call
    /// [`SourceBackedSlideOrderEdit::move_slide`] and
    /// [`SourceBackedSlideOrderEdit::swap_slides`] stage against the order
    /// produced by the preceding operations. The same position is an
    /// explicit exact no-op.
    #[cfg(test)]
    pub(crate) fn edit_slide_order(&self) -> Result<SourceBackedSlideOrderEdit> {
        let source = capture(self)?;
        Ok(SourceBackedSlideOrderEdit {
            working: source.xml.clone(),
            bindings: source.bindings.clone(),
            source,
        })
    }

    /// Publish an exact-source-checked slide-order commit to a sequential
    /// stream.
    #[cfg(test)]
    pub(crate) fn publish_slide_order_commit_to_stream<W: Write>(
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
    #[cfg(test)]
    pub(crate) fn publish_slide_order_patch_to_stream<W: Write>(
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

#[cfg(test)]
impl SourceBackedSlideOrderSnapshot {
    /// Whether this snapshot is exactly the same source snapshot as `other`.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn is_exactly_same_as(&self, other: &Self) -> bool {
        self.same_source(other)
    }

    #[cfg(test)]
    fn check_execution(&self) -> Result<()> {
        check_execution_context(self.context.as_ref())
    }

    #[cfg(test)]
    fn same_source(&self, other: &Self) -> bool {
        self.presentation_uri == other.presentation_uri
            && self.xml.as_bytes() == other.xml.as_bytes()
            && self.bindings == other.bindings
            && self.relationships == other.relationships
            && self.source_version == other.source_version
            && self.lineage == other.lineage
    }
}

#[cfg(test)]
impl SourceBackedSlideOrderEdit {
    /// Stage a positional move in the currently staged slide list.
    ///
    /// `from == to` is a successful, explicit exact no-op. Every candidate is
    /// rebuilt from the original source XML and the staged binding
    /// permutation, so cancelling a sequence restores the original bytes and
    /// lexical entry spans exactly.
    #[cfg(test)]
    pub(crate) fn move_slide(&mut self, from: Position, to: Position) -> Result<bool> {
        self.source.check_execution()?;
        let length = self.source.bindings.len();
        let from = from.get();
        let to = to.get();
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
            return Ok(false);
        }

        let mut ordered = clone_bindings(&self.bindings)?;
        let binding = ordered.remove(from);
        ordered.insert(to, binding);
        self.stage_order(ordered)
    }

    /// Stage an exchange of two positions in the currently staged slide list.
    ///
    /// `left == right` is a successful, explicit exact no-op. The operation
    /// is failure-atomic and composes with preceding moves and swaps.
    #[cfg(test)]
    pub(crate) fn swap_slides(&mut self, left: Position, right: Position) -> Result<bool> {
        self.source.check_execution()?;
        let length = self.source.bindings.len();
        let left = left.get();
        let right = right.get();
        if left >= length {
            return Err(Error::SlideIndexOutOfBounds {
                index: left,
                len: length,
            });
        }
        if right >= length {
            return Err(Error::SlideIndexOutOfBounds {
                index: right,
                len: length,
            });
        }
        if left == right {
            self.source.check_execution()?;
            return Ok(false);
        }

        let mut ordered = clone_bindings(&self.bindings)?;
        ordered.swap(left, right);
        self.stage_order(ordered)
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    #[cfg(test)]
    pub(crate) fn commit(self) -> Result<SourceBackedSlideOrderCommit> {
        self.source.check_execution()?;
        Ok(self.into_commit())
    }

    /// Rebuild the candidate from the immutable source's lexical spans.
    ///
    /// This keeps all inter-entry trivia in its original slot while applying
    /// the currently staged binding permutation. It also avoids retaining a
    /// chain of intermediate XML buffers when a caller composes operations.
    #[cfg(test)]
    fn stage_order(&mut self, ordered: Vec<SlideOrderBinding>) -> Result<bool> {
        self.source.check_execution()?;
        if ordered.as_slice() == self.bindings.as_ref() {
            return Ok(false);
        }
        validate_for_rewrite(&self.source)?;
        let current_refs = binding_refs(&self.source.bindings)?;
        let ordered_refs = binding_refs(&ordered)?;
        let candidate_working = if ordered.as_slice() == self.source.bindings.as_ref() {
            self.source.xml.clone()
        } else {
            let reservation = reserve_order_output(&self.source)?;
            let xml = crate::opened::reorder_slide_bindings(
                self.source.xml.as_bytes(),
                &current_refs,
                &ordered_refs,
                self.source.context.as_ref(),
            )?;
            match reservation {
                Some(reservation) => OrderPayload::ReservedEdit {
                    bytes: Arc::new(xml),
                    _reservation: reservation,
                },
                None => OrderPayload::Edited(Arc::new(xml)),
            }
        };
        let candidate_bindings = ordered.into_boxed_slice();
        self.source.check_execution()?;
        self.working = candidate_working;
        self.bindings = candidate_bindings;
        Ok(true)
    }

    #[cfg(test)]
    fn into_commit(self) -> SourceBackedSlideOrderCommit {
        let snapshot = SourceBackedSlideOrderSnapshot {
            presentation_uri: self.source.presentation_uri.clone(),
            xml: self.working,
            bindings: self.bindings,
            relationships: self.source.relationships.clone(),
            source_version: self.source.source_version,
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

#[cfg(test)]
impl SourceBackedSlideOrderPatch {
    /// Exact immutable source required by this patch.
    #[must_use]
    #[cfg(test)]
    pub(crate) const fn source(&self) -> &SourceBackedSlideOrderSnapshot {
        &self.before
    }

    #[cfg(test)]
    fn is_changed(&self) -> bool {
        !self.before.same_source(&self.after)
    }

    /// Return the exact source-bound inverse patch.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply only to the exact source captured by this patch.
    #[cfg(test)]
    pub(crate) fn apply(
        &self,
        source: &SourceBackedSlideOrderSnapshot,
    ) -> Result<SourceBackedSlideOrderSnapshot> {
        self.before.check_execution()?;
        source.check_execution()?;
        if !source.same_source(&self.before) {
            return Err(Error::PatchConflict);
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

#[cfg(test)]
impl SourceBackedSlideOrderCommit {
    /// Candidate snapshot after this edit.
    #[must_use]
    #[cfg(test)]
    pub(crate) const fn snapshot(&self) -> &SourceBackedSlideOrderSnapshot {
        &self.snapshot
    }

    /// Exact reversible patch for this edit.
    #[must_use]
    #[cfg(test)]
    pub(crate) const fn patch(&self) -> &SourceBackedSlideOrderPatch {
        &self.patch
    }

    #[cfg(test)]
    fn is_changed(&self) -> bool {
        self.patch.is_changed()
    }

    /// The outcome of this edit or commit.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn diagnostics(&self) -> crate::diagnostics::Diagnostics {
        crate::diagnostics::Diagnostics::new(self.is_changed())
    }

    /// Exact reversible inverse patch for this edit.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn inverse(&self) -> SourceBackedSlideOrderPatch {
        self.patch.inverse()
    }
}

#[cfg(test)]
fn capture(editor: &SourceBackedPresentationEditor) -> Result<SourceBackedSlideOrderSnapshot> {
    editor.package.check_execution()?;
    let source_version = editor.package.source_version()?;
    let view = editor.package.main_document_part()?;
    let raw = editor._presentation.source_data().clone();
    editor.package.check_execution()?;
    if view.partname() != editor._presentation.source_partname() {
        return Err(Error::PatchConflict);
    }

    let bindings = slide_order_bindings(editor)?;
    let relationships = root_relationship_bindings(view.rels())?;
    let pinned_relationships =
        root_relationship_bindings(editor._presentation.source_relationships())?;
    if relationships != pinned_relationships {
        return Err(Error::PatchConflict);
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

#[cfg(test)]
fn validate_for_rewrite(snapshot: &SourceBackedSlideOrderSnapshot) -> Result<()> {
    snapshot.check_execution()?;
    reject_root_features(snapshot.xml.as_bytes(), snapshot.context.as_ref())?;
    // This standalone editor rebuilds each candidate from the source, so it
    // cannot carry the section placement of a sequence of moves; the
    // primary source editor maintains sections (record 0303).
    if crate::opened::sections::present(snapshot.xml.as_bytes())
        && crate::opened::sections::scan(snapshot.xml.as_bytes())?.is_some()
    {
        return Err(refusal(
            SlideOrderRefusal::Sections,
            "presentation section metadata owns slide positions",
        ));
    }
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

/// Validate a primary source-order candidate before it is retained in a
/// package revision.  The full root XML is scanned and its complete stable
/// binding list is checked; the semantic metadata owners are admitted through
/// the same publication boundary used by the standalone order publisher.
pub(super) fn validate_source_order_root(
    package: &litchi_opc::SourceBackedPackage,
    presentation_uri: &PackURI,
    xml: &[u8],
    order: &[StableSlideKey],
    context: Option<&ExecutionContext>,
) -> Result<()> {
    reject_root_features(xml, context)?;
    let bindings = order
        .iter()
        .map(|key| (key.slide_id, key.relationship_id.as_str()))
        .collect::<Vec<_>>();
    crate::opened::validate_slide_bindings(xml, &bindings, context).map_err(
        |error| match error {
            Error::Invalid(detail) => refusal(SlideOrderRefusal::AmbiguousBindings, detail),
            other => other,
        },
    )?;
    validate_effective_publication_source(package, presentation_uri)
}

/// Validate the retained package boundary before a primary root/order
/// replacement is written alongside slide overlays.
pub(super) fn validate_source_order_publication(
    package: &litchi_opc::SourceBackedPackage,
    presentation_uri: &PackURI,
) -> Result<()> {
    validate_effective_publication_source(package, presentation_uri)
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
    let view_limits = crate::view_properties::Limits::from_source_limits(package.read_limits());
    let presentation_limits =
        crate::presentation_properties::Limits::from_source_limits(package.read_limits());
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
        match owner {
            MetadataOwner::ViewProperties => {
                inspect_metadata_xml(data.as_bytes(), owner, context.as_ref())?;
            },
            MetadataOwner::PresentationProperties => {
                inspect_presentation_properties_xml(data.as_bytes(), context.as_ref())?;
            },
        }
        check_execution_context(context.as_ref())?;
        match owner {
            MetadataOwner::ViewProperties => {
                let properties = crate::view_properties::ViewProperties::parse_strict(
                    data.as_bytes(),
                    view_limits,
                )
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
                let properties = crate::presentation_properties::Properties::parse_strict(
                    data.as_bytes(),
                    presentation_limits,
                )
                .map_err(|error| {
                    metadata_refusal(
                        owner,
                        format!("presentation properties parse failed: {error}"),
                    )
                })?;
                validate_modeled_presentation_properties(&properties)?;
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
    // Names are resolved below on namespace values as written; a declaration
    // that spells the markup-compatibility namespace with references binds it
    // all the same (record 0011).
    if litchi_ooxml_common::mce::binds_namespace_by_reference(xml) {
        return Err(metadata_refusal(
            owner,
            "positional metadata contains markup-compatibility content",
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut nodes = 0usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let resolver = reader.resolver();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) | Event::Empty(element) => {
                bump_xml(&mut nodes)?;
                if nodes % 256 == 0 {
                    check_execution_context(context)?;
                }
                if is_mce_namespace(&namespace) || has_mce_attribute(&element, resolver)? {
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

/// Scan `presProps.xml` before projecting its typed model.
///
/// Presentation properties have a small modeled extension vocabulary.  The
/// root and slide-show extension lists are the only places where `p:extLst`
/// and `p:ext` may occur; a same-named element in an opaque DrawingML color
/// or an extension payload is still an unsupported owner.  The typed parser
/// below classifies the permitted extension URIs and rejects unknown values.
fn inspect_presentation_properties_xml(
    xml: &[u8],
    context: Option<&ExecutionContext>,
) -> Result<()> {
    // Names are resolved below on namespace values as written; a declaration
    // that spells the markup-compatibility namespace with references binds it
    // all the same (record 0011).
    if litchi_ooxml_common::mce::binds_namespace_by_reference(xml) {
        return Err(metadata_refusal(
            MetadataOwner::PresentationProperties,
            "positional metadata contains markup-compatibility content",
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack = Vec::new();
    let mut nodes = 0usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let resolver = reader.resolver();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                bump_xml(&mut nodes)?;
                if nodes % 256 == 0 {
                    check_execution_context(context)?;
                }
                if stack.len() >= MAX_XML_DEPTH {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlDepth,
                        ),
                        observed: stack.len().checked_add(1).unwrap_or(usize::MAX),
                        maximum: MAX_XML_DEPTH,
                    });
                }
                reject_mce_element(&namespace, &element, resolver).map_err(|error| {
                    metadata_error(MetadataOwner::PresentationProperties, error)
                })?;
                validate_presentation_extension_marker(
                    &stack,
                    &namespace,
                    element.name().local_name().as_ref(),
                )?;
                validate_color_content(&stack, &namespace, &element, resolver)?;
                stack.push(metadata_element(
                    &namespace,
                    element.name().local_name().as_ref(),
                ));
            },
            Event::Empty(element) => {
                bump_xml(&mut nodes)?;
                if nodes % 256 == 0 {
                    check_execution_context(context)?;
                }
                reject_mce_element(&namespace, &element, resolver).map_err(|error| {
                    metadata_error(MetadataOwner::PresentationProperties, error)
                })?;
                validate_presentation_extension_marker(
                    &stack,
                    &namespace,
                    element.name().local_name().as_ref(),
                )?;
                validate_color_content(&stack, &namespace, &element, resolver)?;
            },
            Event::End(_) => {
                if stack.pop().is_none() {
                    return Err(metadata_refusal(
                        MetadataOwner::PresentationProperties,
                        "presentation properties contain an unexpected closing tag",
                    ));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(metadata_refusal(
                    MetadataOwner::PresentationProperties,
                    "positional metadata contains a DTD or processing instruction",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(metadata_refusal(
            MetadataOwner::PresentationProperties,
            "presentation properties contain an unclosed element",
        ));
    }
    check_execution_context(context)?;
    Ok(())
}

fn validate_modeled_presentation_properties(
    properties: &crate::presentation_properties::Properties,
) -> Result<()> {
    if properties.extensions.iter().any(|extension| {
        matches!(
            extension,
            crate::presentation_properties::Extension::Unknown(_)
        )
    }) || properties.show.as_ref().is_some_and(|show| {
        show.extensions.iter().any(|extension| {
            matches!(
                extension,
                crate::presentation_properties::ShowExtension::Unknown(_)
            )
        })
    }) {
        return Err(metadata_refusal(
            MetadataOwner::PresentationProperties,
            "presentation properties contain an unknown extension",
        ));
    }
    if properties
        .show
        .as_ref()
        .is_some_and(|show| show.pen_color.as_ref().is_some_and(unsupported_color))
        || properties.show.as_ref().is_some_and(|show| {
            show.extensions.iter().any(|extension| {
                let crate::presentation_properties::ShowExtension::LaserColor(color) = extension
                else {
                    return false;
                };
                unsupported_color(color)
            })
        })
        || properties.recent_colors.iter().any(unsupported_color)
    {
        return Err(metadata_refusal(
            MetadataOwner::PresentationProperties,
            "presentation properties contain an unsupported DrawingML color payload",
        ));
    }
    if properties
        .html_publish
        .as_ref()
        .is_some_and(|html| html.extension_xml.is_some())
        || properties
            .web
            .as_ref()
            .is_some_and(|web| web.extension_xml.is_some())
        || properties
            .print
            .as_ref()
            .is_some_and(|print| print.extension_xml.is_some())
    {
        return Err(metadata_refusal(
            MetadataOwner::PresentationProperties,
            "presentation properties contain opaque extension content",
        ));
    }
    Ok(())
}

fn unsupported_color(color: &crate::presentation_properties::Color) -> bool {
    match litchi_drawingml::color::read(&color.xml, &litchi_drawingml::color::Limits::DEFAULT) {
        Ok(value) => value.is_unknown(),
        Err(_) => true,
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum MetadataElement {
    Presentation,
    Show,
    ExtensionList,
    Extension,
    LaserColor,
    PenColor,
    Color,
    Other,
}

fn metadata_element(namespace: &ResolveResult<'_>, local_name: &[u8]) -> MetadataElement {
    if local_name == b"laserClr"
        && matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if *value == POWERPOINT_2010_NAMESPACE
        )
    {
        return MetadataElement::LaserColor;
    }
    if is_drawingml_namespace(namespace)
        && matches!(
            local_name,
            b"scrgbClr" | b"srgbClr" | b"hslClr" | b"sysClr" | b"schemeClr" | b"prstClr"
        )
    {
        return MetadataElement::Color;
    }
    if !is_presentation_namespace(namespace) {
        return MetadataElement::Other;
    }
    match local_name {
        b"presentationPr" => MetadataElement::Presentation,
        b"showPr" => MetadataElement::Show,
        b"penClr" => MetadataElement::PenColor,
        b"extLst" => MetadataElement::ExtensionList,
        b"ext" => MetadataElement::Extension,
        _ => MetadataElement::Other,
    }
}

fn validate_color_content(
    stack: &[MetadataElement],
    namespace: &ResolveResult<'_>,
    element: &quick_xml::events::BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
) -> Result<()> {
    let local_name = element.name().local_name();
    let current = metadata_element(namespace, local_name.as_ref());
    if matches!(
        current,
        MetadataElement::LaserColor | MetadataElement::PenColor
    ) {
        return reject_wrapper_attributes(
            element,
            "presentation properties contain attributes on an opaque color wrapper",
        );
    }
    let inside_color = stack.contains(&MetadataElement::Color);
    let inside_color_wrapper =
        stack.contains(&MetadataElement::LaserColor) || stack.contains(&MetadataElement::PenColor);
    if !matches!(current, MetadataElement::Color) && !inside_color && !inside_color_wrapper {
        return Ok(());
    }
    if !is_drawingml_namespace(namespace) {
        return Err(metadata_refusal(
            MetadataOwner::PresentationProperties,
            "presentation properties contain a foreign XML namespace in a DrawingML color",
        ));
    }
    reject_non_namespace_attributes(
        element,
        resolver,
        "presentation properties contain a foreign or relationship attribute in a DrawingML color",
    )
}

fn reject_wrapper_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    detail: &str,
) -> Result<()> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_namespace_binding().is_none() {
            return Err(metadata_refusal(
                MetadataOwner::PresentationProperties,
                detail,
            ));
        }
    }
    Ok(())
}

fn reject_non_namespace_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    detail: &str,
) -> Result<()> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) {
            return Err(metadata_refusal(
                MetadataOwner::PresentationProperties,
                detail,
            ));
        }
    }
    Ok(())
}

fn is_drawingml_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE
    )
}

fn validate_presentation_extension_marker(
    stack: &[MetadataElement],
    namespace: &ResolveResult<'_>,
    local_name: &[u8],
) -> Result<()> {
    if !matches!(local_name, b"extLst" | b"ext") {
        return Ok(());
    }
    let allowed = is_presentation_namespace(namespace)
        && match local_name {
            b"extLst" => matches!(
                stack,
                [MetadataElement::Presentation]
                    | [MetadataElement::Presentation, MetadataElement::Show]
            ),
            b"ext" => matches!(
                stack,
                [
                    MetadataElement::Presentation,
                    MetadataElement::ExtensionList
                ] | [
                    MetadataElement::Presentation,
                    MetadataElement::Show,
                    MetadataElement::ExtensionList,
                ]
            ),
            _ => false,
        };
    if allowed {
        return Ok(());
    }
    Err(metadata_refusal(
        MetadataOwner::PresentationProperties,
        "presentation properties contain an extension marker outside a modeled extension list",
    ))
}

fn metadata_error(_owner: MetadataOwner, error: Error) -> Error {
    error
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

#[cfg(test)]
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

#[cfg(test)]
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

#[cfg(test)]
fn binding_refs(bindings: &[SlideOrderBinding]) -> Result<Vec<(u32, &str)>> {
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

#[cfg(test)]
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
    // Whether the root's `p:extLst` is open: its `p:ext` children must name
    // non-positional extensions or the section list (record 0303), whose
    // membership slide order and removal maintain.
    let mut extension_list = false;
    let mut section_extension = false;
    let mut sections = false;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let resolver = reader.resolver();
        let (namespace, event) = resolver.resolve_event(event);
        if let Event::Start(element) | Event::Empty(element) = &event {
            if depth == 1 && element.name().local_name().as_ref() == b"extLst" {
                extension_list = matches!(event, Event::Start(_));
            } else if depth == 2 && extension_list {
                let uri = litchi_ooxml_common::xml::unqualified_attribute_value(
                    element,
                    b"uri",
                    reader.decoder(),
                )?;
                let section = uri.as_deref() == Some(crate::opened::SECTION_EXTENSION_URI);
                if !is_presentation_namespace(&namespace)
                    || element.name().local_name().as_ref() != b"ext"
                    || !(section
                        || uri.as_deref().is_some_and(|uri| {
                            crate::opened::NON_POSITIONAL_EXTENSIONS.contains(&uri)
                        }))
                {
                    return Err(refusal(
                        SlideOrderRefusal::PositionalOwner,
                        "presentation root extension may own slide positions",
                    ));
                }
                if section {
                    sections = true;
                    section_extension = matches!(event, Event::Start(_));
                }
            }
        }
        if matches!(event, Event::End(_)) && depth == 2 {
            extension_list = false;
        }
        if matches!(event, Event::End(_)) && depth == 3 {
            section_extension = false;
        }
        match event {
            Event::Start(element) => {
                bump_xml(&mut nodes)?;
                if nodes % 256 == 0 {
                    check_execution_context(context)?;
                }
                if depth >= MAX_XML_DEPTH {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlDepth,
                        ),
                        observed: depth.checked_add(1).unwrap_or(usize::MAX),
                        maximum: MAX_XML_DEPTH,
                    });
                }
                reject_mce_element(&namespace, &element, resolver)?;
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
                if !section_extension {
                    reject_positional_element(element.name().local_name().as_ref())?;
                }
                depth += 1;
            },
            Event::Empty(element) => {
                bump_xml(&mut nodes)?;
                if nodes % 256 == 0 {
                    check_execution_context(context)?;
                }
                reject_mce_element(&namespace, &element, resolver)?;
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
                if !section_extension {
                    reject_positional_element(element.name().local_name().as_ref())?;
                }
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
    if sections {
        // The section list must read exactly; an unreadable one is refused.
        crate::opened::sections::scan(xml)?;
    }
    Ok(())
}

fn reject_mce_element(
    namespace: &ResolveResult<'_>,
    element: &quick_xml::events::BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
) -> Result<()> {
    if is_mce_namespace(namespace) {
        return Err(refusal(
            SlideOrderRefusal::MarkupCompatibility,
            "presentation XML contains markup-compatibility content",
        ));
    }
    if has_mce_attribute(element, resolver)? {
        return Err(refusal(
            SlideOrderRefusal::MarkupCompatibility,
            "presentation XML contains markup-compatibility attributes",
        ));
    }
    Ok(())
}

fn has_mce_attribute(
    element: &quick_xml::events::BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
) -> Result<bool> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if is_mce_namespace(&namespace)
            && matches!(
                local.as_ref(),
                b"Ignorable"
                    | b"MustUnderstand"
                    | b"ProcessContent"
                    | b"PreserveAttributes"
                    | b"PreserveElements"
            )
        {
            return Ok(true);
        }
    }
    Ok(false)
}

fn validate_root_attributes(element: &quick_xml::events::BytesStart<'_>) -> Result<()> {
    for attribute in element.checked_attributes() {
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
        b"photoAlbum" => Err(refusal(
            SlideOrderRefusal::PositionalOwner,
            "presentation photo-album metadata may own slide positions",
        )),
        // Each extension is admitted by name in `reject_root_features`.
        b"extLst" => Ok(()),
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
            if *value == MCE_NAMESPACE
    )
}

fn check_execution_context(context: Option<&ExecutionContext>) -> Result<()> {
    if let Some(context) = context {
        context.check().map_err(|error| {
            Error::Opc(match error {
                ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                error => litchi_opc::OpcError::Execution(error),
            })
        })?;
    }
    Ok(())
}

/// Admit the exact output allocation required by a slide-order rewrite.
///
/// Reordering only permutes existing entry spans, so the candidate has the
/// same byte length as the retained presentation root. The reservation is
/// attached to the candidate payload and therefore follows clones of the
/// snapshot and patch until the last retaining snapshot or patch is dropped.
#[cfg(test)]
fn reserve_order_output(
    snapshot: &SourceBackedSlideOrderSnapshot,
) -> Result<Option<Arc<Reservation>>> {
    reserve_managed_output(snapshot.context.as_ref(), snapshot.xml.as_bytes().len())
}

/// Reserve an edit-owned output before the lowering step allocates its
/// replacement bytes.  The returned token is intentionally shareable so the
/// reservation follows immutable candidate snapshots, patches, and inverse
/// patches until their last owner is dropped.
pub(super) fn reserve_managed_output(
    context: Option<&ExecutionContext>,
    bytes: usize,
) -> Result<Option<Arc<Reservation>>> {
    let Some(context) = context else {
        return Ok(None);
    };
    let bytes = u64::try_from(bytes)
        .map_err(|_| Error::Invalid("source-backed output exceeds u64".into()))?;
    context
        .reserve(Resource::Memory, bytes)
        .map(Arc::new)
        .map(Some)
        .map_err(|error| {
            Error::Opc(match error {
                ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                error => litchi_opc::OpcError::Execution(error),
            })
        })
}

/// Lower an order permutation identified by stable physical slide bindings.
///
/// The native XML rewriter retains source lexical spans and inter-entry trivia;
/// this adapter only translates stable keys to its checked `(id, r:id)` input.
/// It validates that both lists are complete, unique permutations before any
/// candidate bytes are allocated.
pub(super) fn reorder_stable_bindings(
    xml: &[u8],
    current: &[StableSlideKey],
    ordered: &[StableSlideKey],
    context: Option<&ExecutionContext>,
) -> Result<Vec<u8>> {
    if current.len() != ordered.len() {
        return Err(refusal(
            SlideOrderRefusal::AmbiguousBindings,
            "source slide order is not a complete stable-binding permutation",
        ));
    }
    let mut current_refs = Vec::new();
    current_refs
        .try_reserve_exact(current.len())
        .map_err(allocation)?;
    let mut current_keys = std::collections::HashSet::new();
    current_keys
        .try_reserve(current.len())
        .map_err(allocation)?;
    for key in current {
        current_keys.insert(key);
        current_refs.push((key.slide_id, key.relationship_id.as_str()));
    }
    let mut ordered_refs = Vec::new();
    ordered_refs
        .try_reserve_exact(ordered.len())
        .map_err(allocation)?;
    let mut seen = std::collections::HashSet::new();
    seen.try_reserve(ordered.len()).map_err(allocation)?;
    for key in ordered {
        if !seen.insert(key) {
            return Err(refusal(
                SlideOrderRefusal::AmbiguousBindings,
                "source slide order repeats a stable binding",
            ));
        }
        if !current_keys.contains(key) {
            return Err(refusal(
                SlideOrderRefusal::AmbiguousBindings,
                "source slide order references an unknown stable binding",
            ));
        }
        ordered_refs.push((key.slide_id, key.relationship_id.as_str()));
    }
    if seen.len() != current.len() {
        return Err(refusal(
            SlideOrderRefusal::AmbiguousBindings,
            "source slide order omits a stable binding",
        ));
    }
    crate::opened::reorder_slide_bindings(xml, &current_refs, &ordered_refs, context)
}

/// Produce a deterministic root representation for a source-owner digest.
///
/// The ordered slide entries are sorted by their physical stable identity
/// before the native span permutation runs.  This keeps the digest invariant
/// under a valid order change while retaining every other root byte and
/// lexical fragment, so a durable patch cannot rebind itself to a different
/// presentation root by changing both root preimages.
pub(super) fn canonical_source_order_xml(
    xml: &[u8],
    order: &[StableSlideKey],
    context: Option<&ExecutionContext>,
) -> Result<Vec<u8>> {
    let mut sorted = order.to_vec();
    sorted.sort_unstable_by(|left, right| {
        left.part_uri
            .as_str()
            .cmp(right.part_uri.as_str())
            .then(left.slide_id.cmp(&right.slide_id))
            .then(left.relationship_id.cmp(&right.relationship_id))
    });
    let canonical = reorder_stable_bindings(xml, order, &sorted, context)?;
    // Section membership follows the order (record 0303); the exact root
    // preimages guard it, so the order-invariant digest leaves it out.
    if crate::opened::sections::present(&canonical) {
        if let Some((stripped, _)) = crate::opened::sections::strip(&canonical)? {
            return Ok(stripped);
        }
    }
    Ok(canonical)
}

fn allocation(source: std::collections::TryReserveError) -> Error {
    Error::Allocation {
        resource: "source-backed stable slide order",
        source,
    }
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
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlNodes),
            observed: *nodes,
            maximum: MAX_XML_NODES,
        });
    }
    Ok(())
}

fn refusal(kind: SlideOrderRefusal, _detail: impl Into<String>) -> Error {
    Error::slide_order_refused(kind)
}

#[cfg(test)]
mod tests {
    use super::*;

    const FOREIGN_MCE: &[u8] = b"http://purl.oclc.org/ooxml/markup-compatibility/2006";

    /// Record 0011: a declaration that spells the namespace with references
    /// binds it for preprocessing, so positional metadata refuses it as it
    /// refuses a literal one.
    #[test]
    fn escaped_markup_compatibility_is_refused_in_positional_metadata() {
        let xml = format!(
            r#"<p:viewPr xmlns:p="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/200&#54;" mc:Ignorable="p"/>"#,
            std::str::from_utf8(PRESENTATIONML_NAMESPACE).unwrap(),
        );
        assert_refused_kind!(
            inspect_metadata_xml(xml.as_bytes(), MetadataOwner::ViewProperties, None),
            SlideOrderRefusal::ViewProperties.kind()
        );
    }

    #[test]
    fn only_canonical_mce_attributes_are_classified_as_markup_compatibility() {
        let canonical_root = format!(
            r#"<p:presentation xmlns:p="{}" xmlns:mc="{}" mc:Ignorable="p"/>"#,
            std::str::from_utf8(PRESENTATIONML_NAMESPACE).unwrap(),
            std::str::from_utf8(MCE_NAMESPACE).unwrap(),
        );
        assert_refused_kind!(
            reject_root_features(canonical_root.as_bytes(), None),
            SlideOrderRefusal::MarkupCompatibility.kind()
        );

        let foreign_root = format!(
            r#"<p:presentation xmlns:p="{}" xmlns:foreign="{}" foreign:Ignorable="p"/>"#,
            std::str::from_utf8(PRESENTATIONML_NAMESPACE).unwrap(),
            std::str::from_utf8(FOREIGN_MCE).unwrap(),
        );
        assert_refused_kind!(
            reject_root_features(foreign_root.as_bytes(), None),
            SlideOrderRefusal::PositionalOwner.kind()
        );

        let canonical_metadata = format!(
            r#"<p:viewPr xmlns:p="{}" xmlns:mc="{}" mc:Ignorable="p"/>"#,
            std::str::from_utf8(PRESENTATIONML_NAMESPACE).unwrap(),
            std::str::from_utf8(MCE_NAMESPACE).unwrap(),
        );
        assert_refused_kind!(
            inspect_metadata_xml(
                canonical_metadata.as_bytes(),
                MetadataOwner::ViewProperties,
                None,
            ),
            SlideOrderRefusal::ViewProperties.kind()
        );

        let foreign_metadata = format!(
            r#"<p:viewPr xmlns:p="{}" xmlns:foreign="{}" foreign:Ignorable="p"/>"#,
            std::str::from_utf8(PRESENTATIONML_NAMESPACE).unwrap(),
            std::str::from_utf8(FOREIGN_MCE).unwrap(),
        );
        assert!(
            inspect_metadata_xml(
                foreign_metadata.as_bytes(),
                MetadataOwner::ViewProperties,
                None,
            )
            .is_ok()
        );
    }

    #[test]
    fn modeled_presentation_extensions_are_admitted_only_at_owner_lists() {
        let presentation = std::str::from_utf8(PRESENTATIONML_NAMESPACE).unwrap();
        let p14 = std::str::from_utf8(POWERPOINT_2010_NAMESPACE).unwrap();
        let drawing = std::str::from_utf8(DRAWINGML_NAMESPACE).unwrap();
        let root = format!(
            r#"<p:presentationPr xmlns:p="{presentation}" xmlns:p14="{p14}"><p:showPr><p:penClr><a:srgbClr xmlns:a="{drawing}" val="112233"/></p:penClr><p:extLst><p:ext uri="{show_uri}"><p14:showMediaCtrls val="0"/></p:ext></p:extLst></p:showPr><p:extLst><p:ext uri="{uri}"><p14:discardImageEditData val="0"/></p:ext></p:extLst></p:presentationPr>"#,
            uri = crate::presentation_properties::DISCARD_IMAGE_EDIT_DATA_URI,
            show_uri = crate::presentation_properties::SHOW_MEDIA_CONTROLS_URI,
        );
        assert!(inspect_presentation_properties_xml(root.as_bytes(), None).is_ok());
        let properties = crate::presentation_properties::Properties::parse_strict(
            root.as_bytes(),
            crate::presentation_properties::Limits::default(),
        )
        .unwrap();
        assert!(validate_modeled_presentation_properties(&properties).is_ok());
    }

    #[test]
    fn presentation_extension_markers_inside_opaque_color_are_refused() {
        let presentation = std::str::from_utf8(PRESENTATIONML_NAMESPACE).unwrap();
        let drawing = std::str::from_utf8(DRAWINGML_NAMESPACE).unwrap();
        let foreign = format!(
            r#"<p:presentationPr xmlns:p="{presentation}"><p:showPr><p:penClr><a:srgbClr xmlns:a="{drawing}" val="112233"><a:extLst/></a:srgbClr></p:penClr></p:showPr></p:presentationPr>"#
        );
        assert_refused_kind!(
            inspect_presentation_properties_xml(foreign.as_bytes(), None),
            SlideOrderRefusal::PresentationProperties.kind()
        );
        let foreign_namespace = format!(
            r#"<p:presentationPr xmlns:p="{presentation}"><p:showPr><p:penClr><a:srgbClr xmlns:a="{drawing}" xmlns:v="urn:foreign" v:href="https://example.invalid" val="112233"/></p:penClr></p:showPr></p:presentationPr>"#
        );
        assert_refused_kind!(
            inspect_presentation_properties_xml(foreign_namespace.as_bytes(), None),
            SlideOrderRefusal::PresentationProperties.kind()
        );
        let foreign_child = format!(
            r#"<p:presentationPr xmlns:p="{presentation}" xmlns:p14="{p14}" xmlns:a="{drawing}" xmlns:v="urn:foreign"><p:showPr><p:extLst><p:ext uri="{show_uri}"><p14:laserClr><a:srgbClr val="112233"><v:payload/></a:srgbClr></p14:laserClr></p:ext></p:extLst></p:showPr></p:presentationPr>"#,
            p14 = std::str::from_utf8(POWERPOINT_2010_NAMESPACE).unwrap(),
            drawing = std::str::from_utf8(DRAWINGML_NAMESPACE).unwrap(),
            show_uri = crate::presentation_properties::LASER_COLOR_URI,
        );
        assert_refused_kind!(
            inspect_presentation_properties_xml(foreign_child.as_bytes(), None),
            SlideOrderRefusal::PresentationProperties.kind()
        );
        let relationship_wrapper = format!(
            r#"<p:presentationPr xmlns:p="{presentation}" xmlns:p14="{p14}" xmlns:a="{drawing}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:showPr><p:extLst><p:ext uri="{show_uri}"><p14:laserClr r:id="rId2"><a:srgbClr val="112233"/></p14:laserClr></p:ext></p:extLst></p:showPr></p:presentationPr>"#,
            p14 = std::str::from_utf8(POWERPOINT_2010_NAMESPACE).unwrap(),
            drawing = std::str::from_utf8(DRAWINGML_NAMESPACE).unwrap(),
            show_uri = crate::presentation_properties::LASER_COLOR_URI,
        );
        assert_refused_kind!(
            inspect_presentation_properties_xml(relationship_wrapper.as_bytes(), None),
            SlideOrderRefusal::PresentationProperties.kind()
        );
        let foreign_wrapper = format!(
            r#"<p:presentationPr xmlns:p="{presentation}" xmlns:p14="{p14}" xmlns:a="{drawing}" xmlns:v="urn:foreign"><p:showPr><p:extLst><p:ext uri="{show_uri}"><p14:laserClr v:href="https://example.invalid"><a:srgbClr val="112233"/></p14:laserClr></p:ext></p:extLst></p:showPr></p:presentationPr>"#,
            p14 = std::str::from_utf8(POWERPOINT_2010_NAMESPACE).unwrap(),
            drawing = std::str::from_utf8(DRAWINGML_NAMESPACE).unwrap(),
            show_uri = crate::presentation_properties::LASER_COLOR_URI,
        );
        assert_refused_kind!(
            inspect_presentation_properties_xml(foreign_wrapper.as_bytes(), None),
            SlideOrderRefusal::PresentationProperties.kind()
        );
    }

    #[test]
    fn unknown_presentation_extensions_are_refused_after_strict_projection() {
        let presentation = std::str::from_utf8(PRESENTATIONML_NAMESPACE).unwrap();
        let unknown = format!(
            r#"<p:presentationPr xmlns:p="{presentation}" xmlns:x="urn:unknown"><p:extLst><p:ext uri="urn:unknown"><x:payload/></p:ext></p:extLst></p:presentationPr>"#
        );
        inspect_presentation_properties_xml(unknown.as_bytes(), None).unwrap();
        let properties = crate::presentation_properties::Properties::parse_strict(
            unknown.as_bytes(),
            crate::presentation_properties::Limits::default(),
        )
        .unwrap();
        assert_refused_kind!(
            validate_modeled_presentation_properties(&properties),
            SlideOrderRefusal::PresentationProperties.kind()
        );
    }
}
