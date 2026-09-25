//! Source-bound attach/detach transactions for SVG resources on existing
//! PresentationML raster pictures.
//!
//! This owner deliberately models the smallest useful dependency closure: one
//! direct `p:pic`, its compatibility `a:blip`, the optional native SVG
//! extension, the owning slide relationship member, and one SVG media Part.
//! Unknown `a:ext` siblings remain byte-for-byte opaque.  Picture creation,
//! shape reordering, MCE branch selection, linked SVG authoring, and rendering
//! remain outside this capability.

use std::sync::Arc;

use litchi_core::xml::ReaderOrigin;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::reader::{NsReader, Reader};

use litchi_opc::{AuthoredXmlFragment, PackURI, Part, SourceTopologyPlan, TargetMode};

use super::{
    SourceBackedPresentationEditor, SourceBackedSlideSnapshot, SourcePayload, is_png_content_type,
    is_svg_content_type, resolve_picture_target, resolve_svg_target,
    validate_full_slide_picture_relationships, validate_source_slide_root,
};
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

mod owner;
use owner::{ByteRange, ElementRange, PictureLayout};

const RELATIONSHIP_NAMESPACE: &[u8] = litchi_ooxml_common::relationships::TRANSITIONAL_NAMESPACE;
const MAX_GENERATED_NAME_ATTEMPTS: usize = 100_000;

/// A caller-owned SVG payload and optional advanced package identity.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceSvgAttachmentReplacement {
    svg: Arc<Vec<u8>>,
    svg_part_uri: Option<PackURI>,
    relationship_id: Option<String>,
}

impl SourceSvgAttachmentReplacement {
    /// Construct an ordinary attach request.  The package editor allocates a
    /// relationship ID and media URI that do not collide with the source.
    #[must_use]
    pub fn new(svg: impl Into<Vec<u8>>) -> Self {
        Self {
            svg: Arc::new(svg.into()),
            svg_part_uri: None,
            relationship_id: None,
        }
    }

    /// Select an explicit `/ppt/media/` SVG Part URI.  This is an advanced
    /// identity-preserving escape hatch; ordinary callers should omit it.
    #[must_use]
    pub fn with_part_uri(mut self, part_uri: PackURI) -> Self {
        self.svg_part_uri = Some(part_uri);
        self
    }

    /// Select an explicit relationship ID.  It must be unused by the owning
    /// slide relationship member.
    #[must_use]
    pub fn with_relationship_id(mut self, relationship_id: impl Into<String>) -> Self {
        self.relationship_id = Some(relationship_id.into());
        self
    }

    /// Borrow the caller-supplied SVG bytes.
    #[must_use]
    pub fn svg(&self) -> &[u8] {
        self.svg.as_slice()
    }

    /// Borrow the requested advanced SVG Part URI, when supplied.
    #[must_use]
    pub const fn part_uri(&self) -> Option<&PackURI> {
        self.svg_part_uri.as_ref()
    }

    /// Borrow the requested advanced relationship ID, when supplied.
    #[must_use]
    pub fn relationship_id(&self) -> Option<&str> {
        self.relationship_id.as_deref()
    }
}

/// The attached SVG resource in one source-backed lifecycle snapshot.
#[derive(Clone)]
pub struct SourceSvgAttachment {
    relationship_id: String,
    relationship_type: String,
    part_uri: PackURI,
    payload: SourcePayload,
}

impl SourceSvgAttachment {
    /// Relationship ID carried by the native `asvg:svgBlip` element.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Relationship type carried by the owning slide relationship member.
    #[must_use]
    pub fn relationship_type(&self) -> &str {
        &self.relationship_type
    }

    /// Internal SVG media Part URI.
    #[must_use]
    pub const fn part_uri(&self) -> &PackURI {
        &self.part_uri
    }

    /// Borrow the exact source or staged SVG bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.payload.as_bytes()
    }
}

/// Immutable source/target state for an existing raster picture's optional
/// SVG owner.
#[derive(Clone)]
pub struct SourceBackedSvgAttachmentSnapshot {
    pub(super) slide: SourceBackedSlideSnapshot,
    image_position: usize,
    raster_relationship_id: String,
    raster_relationship_type: String,
    raster_part_uri: PackURI,
    svg: Option<SourceSvgAttachment>,
    default_relationship_id: String,
    default_part_uri: PackURI,
    existing_relationship_ids: Arc<[String]>,
    physical_member_names: Arc<[String]>,
    limits: litchi_opc::ReadLimits,
}

/// One-shot source-backed attach/detach edit.
pub struct SourceBackedSvgAttachmentEdit<'a> {
    editor: &'a SourceBackedPresentationEditor,
    source: SourceBackedSvgAttachmentSnapshot,
    slide: SourceBackedSlideSnapshot,
    svg: Option<SourceSvgAttachment>,
    operation_used: bool,
    default_relationship_id: String,
    default_part_uri: PackURI,
}

/// Exact-source reversible patch for an SVG attach/detach operation.
#[derive(Clone)]
pub struct SourceBackedSvgAttachmentPatch {
    before: SourceBackedSvgAttachmentSnapshot,
    after: SourceBackedSvgAttachmentSnapshot,
}

/// Checked attach/detach transaction ready for publication.
pub struct SourceBackedSvgAttachmentCommit {
    snapshot: SourceBackedSvgAttachmentSnapshot,
    patch: SourceBackedSvgAttachmentPatch,
}

impl SourceBackedSvgAttachmentSnapshot {
    /// Zero-based presentation position of the selected slide.
    #[must_use]
    pub const fn slide_position(&self) -> usize {
        self.slide.position()
    }

    /// Zero-based direct-picture position in the selected slide.
    #[must_use]
    pub const fn image_position(&self) -> usize {
        self.image_position
    }

    /// Raster relationship ID retained by the compatibility `a:blip`.
    #[must_use]
    pub fn raster_relationship_id(&self) -> &str {
        &self.raster_relationship_id
    }

    /// Raster relationship type retained by the compatibility `a:blip`.
    #[must_use]
    pub fn raster_relationship_type(&self) -> &str {
        &self.raster_relationship_type
    }

    /// Internal raster fallback Part URI.
    #[must_use]
    pub const fn raster_part_uri(&self) -> &PackURI {
        &self.raster_part_uri
    }

    /// Current optional SVG attachment.
    #[must_use]
    pub fn svg(&self) -> Option<&SourceSvgAttachment> {
        self.svg.as_ref()
    }

    /// Start an editor-authorized edit from this exact snapshot.
    fn edit<'a>(
        &self,
        editor: &'a SourceBackedPresentationEditor,
    ) -> SourceBackedSvgAttachmentEdit<'a> {
        SourceBackedSvgAttachmentEdit {
            editor,
            source: self.clone(),
            slide: self.slide.clone(),
            svg: self.svg.clone(),
            operation_used: false,
            default_relationship_id: self.default_relationship_id.clone(),
            default_part_uri: self.default_part_uri.clone(),
        }
    }

    pub(super) fn same_source(&self, other: &Self) -> bool {
        self.slide.same_source(&other.slide)
            && self.image_position == other.image_position
            && self.raster_relationship_id == other.raster_relationship_id
            && self.raster_relationship_type == other.raster_relationship_type
            && self.raster_part_uri == other.raster_part_uri
            && same_attachment(&self.svg, &other.svg)
            && self.default_relationship_id == other.default_relationship_id
            && self.default_part_uri == other.default_part_uri
            && self.existing_relationship_ids.as_ref() == other.existing_relationship_ids.as_ref()
            && self.physical_member_names.as_ref() == other.physical_member_names.as_ref()
            && self.limits == other.limits
    }
}

impl<'a> SourceBackedSvgAttachmentEdit<'a> {
    /// Exact immutable source snapshot against which this edit was created.
    #[must_use]
    pub const fn source(&self) -> &SourceBackedSvgAttachmentSnapshot {
        &self.source
    }

    /// Attach one SVG using collision-free generated package identities.
    pub fn attach_svg(&mut self, svg: &[u8]) -> Result<bool> {
        let (relationship_id, part_uri) = self.prepare_attachment(svg.len(), None, None)?;
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(svg.len())
            .map_err(|source| Error::Allocation {
                resource: "SVG attachment request",
                source,
            })?;
        owned.extend_from_slice(svg);
        let replacement = SourceSvgAttachmentReplacement {
            svg: Arc::new(owned),
            svg_part_uri: None,
            relationship_id: None,
        };
        self.apply_attachment(relationship_id, part_uri, replacement.svg.clone())
    }

    /// Attach one SVG with optional advanced package identities.
    pub fn attach(&mut self, replacement: &SourceSvgAttachmentReplacement) -> Result<bool> {
        let (relationship_id, part_uri) = self.prepare_attachment(
            replacement.svg.len(),
            replacement.relationship_id.as_deref(),
            replacement.svg_part_uri.as_ref(),
        )?;
        self.apply_attachment(relationship_id, part_uri, Arc::clone(&replacement.svg))
    }

    fn prepare_attachment(
        &self,
        svg_len: usize,
        requested_relationship_id: Option<&str>,
        requested_part_uri: Option<&PackURI>,
    ) -> Result<(String, PackURI)> {
        if self.operation_used {
            return Err(Error::UnsafeEdit {
                operation: "attach_svg",
                reason: "SVG lifecycle edits support one atomic attach or detach",
            });
        }
        if self.svg.is_some() {
            return Err(Error::Relationship(
                "selected raster picture already has an SVG attachment".into(),
            ));
        }
        validate_svg_payload(&self.source.limits, svg_len)?;
        let relationship_id = requested_relationship_id
            .map(ToOwned::to_owned)
            .unwrap_or_else(|| self.default_relationship_id.clone());
        let part_uri = requested_part_uri
            .cloned()
            .unwrap_or_else(|| self.default_part_uri.clone());
        validate_relationship_id(&relationship_id)?;
        if relationship_id == self.source.raster_relationship_id {
            return Err(Error::Relationship(
                "SVG attachment relationship ID collides with the raster relationship".into(),
            ));
        }
        if requested_relationship_id.is_some()
            && self
                .source
                .existing_relationship_ids
                .iter()
                .any(|existing| existing == &relationship_id)
        {
            return Err(Error::Relationship(format!(
                "SVG attachment relationship ID '{relationship_id}' is already in use"
            )));
        }
        super::validate_svg_media_uri(&part_uri)?;
        if requested_part_uri.is_some() {
            let relationship_uri = part_uri
                .rels_uri()
                .map_err(|error| Error::Uri(error.to_string()))?;
            if self.source.physical_member_names.iter().any(|existing| {
                existing.eq_ignore_ascii_case(part_uri.membername())
                    || existing.eq_ignore_ascii_case(relationship_uri.membername())
            }) {
                return Err(Error::Relationship(format!(
                    "SVG attachment target Part '{}' or its relationship member already exists",
                    part_uri.as_str()
                )));
            }
        }
        Ok((relationship_id, part_uri))
    }

    fn apply_attachment(
        &mut self,
        relationship_id: String,
        part_uri: PackURI,
        payload: Arc<Vec<u8>>,
    ) -> Result<bool> {
        let payload = SourcePayload::Edited(payload);
        let svg = SourceSvgAttachment {
            relationship_id: relationship_id.clone(),
            relationship_type: self.source.raster_relationship_type.clone(),
            part_uri: part_uri.clone(),
            payload,
        };
        let layout =
            selected_picture_layout(self.slide.xml.as_bytes(), self.source.image_position)?;
        let rewritten = rewrite_picture_svg(
            self.slide.xml.as_bytes(),
            &layout,
            Some(&relationship_id),
            &self.source.limits,
        )?;
        self.slide.xml = SourcePayload::Edited(Arc::new(rewritten));
        self.svg = Some(svg);
        self.operation_used = true;
        Ok(true)
    }

    /// Detach the selected SVG while retaining the raster fallback.
    pub fn detach(&mut self) -> Result<bool> {
        if self.operation_used {
            return Err(Error::UnsafeEdit {
                operation: "detach_svg",
                reason: "SVG lifecycle edits support one atomic attach or detach",
            });
        }
        let Some(svg) = self.svg.as_ref() else {
            self.operation_used = true;
            return Ok(false);
        };
        let layout =
            selected_picture_layout(self.slide.xml.as_bytes(), self.source.image_position)?;
        let rewritten = rewrite_picture_svg(
            self.slide.xml.as_bytes(),
            &layout,
            None,
            &self.source.limits,
        )?;
        self.slide.xml = SourcePayload::Edited(Arc::new(rewritten));
        let _ = svg;
        self.svg = None;
        self.operation_used = true;
        Ok(true)
    }

    /// Validate the selected dependency closure and freeze the edit into an
    /// exact-source patch. The validation plan is discarded; publication
    /// recaptures the source and rebuilds the physical plan.
    pub fn commit(self) -> Result<SourceBackedSvgAttachmentCommit> {
        self.editor.package.check_execution()?;
        let snapshot = SourceBackedSvgAttachmentSnapshot {
            slide: self.slide,
            image_position: self.source.image_position,
            raster_relationship_id: self.source.raster_relationship_id.clone(),
            raster_relationship_type: self.source.raster_relationship_type.clone(),
            raster_part_uri: self.source.raster_part_uri.clone(),
            svg: self.svg,
            default_relationship_id: self.default_relationship_id,
            default_part_uri: self.default_part_uri,
            existing_relationship_ids: self.source.existing_relationship_ids.clone(),
            physical_member_names: self.source.physical_member_names.clone(),
            limits: self.source.limits,
        };
        let patch = SourceBackedSvgAttachmentPatch {
            before: self.source,
            after: snapshot.clone(),
        };
        if patch.is_changed() {
            let _plan =
                build_attachment_topology_plan(&self.editor.package, &patch.before, &patch.after)?;
        }
        self.editor.package.check_execution()?;
        Ok(SourceBackedSvgAttachmentCommit { snapshot, patch })
    }

    /// Alias for [`Self::commit`], retained for callers that make the
    /// execution check explicit.
    pub fn commit_checked(self) -> Result<SourceBackedSvgAttachmentCommit> {
        self.commit()
    }
}

impl SourceBackedSvgAttachmentPatch {
    /// Exact source snapshot required by this patch.
    #[must_use]
    pub const fn source(&self) -> &SourceBackedSvgAttachmentSnapshot {
        &self.before
    }

    /// Exact target snapshot produced by this patch.
    #[must_use]
    pub const fn target(&self) -> &SourceBackedSvgAttachmentSnapshot {
        &self.after
    }

    /// Whether the selected slide/relationship closure changes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !same_attachment_state(&self.before, &self.after)
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply only to the exact source snapshot captured by this patch.
    pub fn apply(
        &self,
        source: &SourceBackedSvgAttachmentSnapshot,
    ) -> Result<SourceBackedSvgAttachmentSnapshot> {
        self.before.slide.check_execution()?;
        source.slide.check_execution()?;
        if !source.same_source(&self.before) {
            return Err(Error::StaleSource);
        }
        Ok(if self.is_changed() {
            self.after.clone()
        } else {
            source.clone()
        })
    }
}

impl SourceBackedSvgAttachmentCommit {
    /// Candidate target snapshot after this edit.
    #[must_use]
    pub const fn snapshot(&self) -> &SourceBackedSvgAttachmentSnapshot {
        &self.snapshot
    }

    /// Exact-source patch for this edit.
    #[must_use]
    pub const fn patch(&self) -> &SourceBackedSvgAttachmentPatch {
        &self.patch
    }

    /// Whether the selected closure changes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.patch.is_changed()
    }
}

impl SourceBackedPresentationEditor {
    /// Begin an exact-source SVG attach/detach transaction on one direct
    /// raster picture.  Existing linked SVGs and MCE-owned picture branches
    /// are refused before an edit handle is returned.
    pub fn edit_svg_attachment(
        &self,
        slide_position: usize,
        image_position: usize,
    ) -> Result<SourceBackedSvgAttachmentEdit<'_>> {
        self.package.check_execution()?;
        capture_svg_attachment_snapshot(self, slide_position, image_position, "edit_svg_attachment")
            .and_then(|snapshot| {
                snapshot.slide.check_execution()?;
                Ok(snapshot.edit(self))
            })
    }

    /// Publish one exact-source SVG attach/detach commit atomically through a
    /// bounded OPC topology plan.
    pub fn publish_svg_attachment_commit_to_stream<W: std::io::Write>(
        self,
        writer: W,
        commit: &SourceBackedSvgAttachmentCommit,
    ) -> Result<SourceBackedSvgAttachmentSnapshot> {
        self.package.check_execution()?;
        let current = capture_svg_attachment_snapshot(
            &self,
            commit.patch.before.slide.position(),
            commit.patch.before.image_position,
            "publish_svg_attachment_commit_to_stream",
        )?;
        if !current.same_source(&commit.patch.before) {
            return Err(Error::StaleSource);
        }
        let target = commit.patch.apply(&current)?;
        if !target.is_changed_from(&current) {
            self.package
                .write_topology_to_stream(writer, SourceTopologyPlan::new())?;
            return Ok(target);
        }
        let plan = build_attachment_topology_plan(&self.package, &current, &target)?;
        self.package.write_topology_to_stream(writer, plan)?;
        Ok(target)
    }

    /// Compatibility alias with the operation noun used by replacement APIs.
    pub fn publish_svg_lifecycle_commit_to_stream<W: std::io::Write>(
        self,
        writer: W,
        commit: &SourceBackedSvgAttachmentCommit,
    ) -> Result<SourceBackedSvgAttachmentSnapshot> {
        self.publish_svg_attachment_commit_to_stream(writer, commit)
    }
}

impl SourceBackedSvgAttachmentSnapshot {
    fn is_changed_from(&self, other: &Self) -> bool {
        !same_attachment_state(self, other)
    }
}

fn same_attachment_state(
    left: &SourceBackedSvgAttachmentSnapshot,
    right: &SourceBackedSvgAttachmentSnapshot,
) -> bool {
    left.slide.xml.as_bytes() == right.slide.xml.as_bytes()
        && left.image_position == right.image_position
        && left.raster_relationship_id == right.raster_relationship_id
        && left.raster_relationship_type == right.raster_relationship_type
        && left.raster_part_uri == right.raster_part_uri
        && same_attachment(&left.svg, &right.svg)
}

fn same_attachment(
    left: &Option<SourceSvgAttachment>,
    right: &Option<SourceSvgAttachment>,
) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => {
            left.relationship_id == right.relationship_id
                && left.relationship_type == right.relationship_type
                && left.part_uri == right.part_uri
                && left.payload.as_bytes() == right.payload.as_bytes()
        },
        _ => false,
    }
}

fn validate_svg_payload(limits: &litchi_opc::ReadLimits, bytes: usize) -> Result<()> {
    if bytes == 0 {
        return Err(Error::Invalid("SVG payload cannot be empty".into()));
    }
    if bytes as u64 > limits.max_part_bytes() {
        return Err(Error::Limit {
            resource: "SVG attachment Part bytes",
            limit: usize::try_from(limits.max_part_bytes()).unwrap_or(usize::MAX),
        });
    }
    Ok(())
}

fn validate_relationship_id(id: &str) -> Result<()> {
    if id.is_empty()
        || id.len() > litchi_drawingml::svg_blip::MAX_RELATIONSHIP_ID_BYTES
        || !litchi_ooxml_common::xml::is_ncname(id)
    {
        return Err(Error::Relationship(format!(
            "invalid SVG attachment relationship ID '{id}'"
        )));
    }
    Ok(())
}

fn selected_picture_layout(xml: &[u8], image_position: usize) -> Result<PictureLayout> {
    owner::locate(xml, image_position)
}

/// Parse a source-sliced picture with its inherited namespace context. The
/// raw owner scanner is required even when the sliced fragment appears
/// self-contained: the legacy parser has prefix fallbacks which can otherwise
/// silently accept a rebinding different from the full slide's namespace
/// environment.
fn parse_picture_relationship_with_layout(
    layout: &PictureLayout,
    picture_xml: &[u8],
    max_output_bytes: u64,
) -> Result<super::PictureRelationship> {
    let complete = owner::namespace_complete_picture(layout, picture_xml, max_output_bytes)?;
    super::parse_picture_relationship_with_limit(&complete, max_output_bytes)
}

pub(super) fn parse_picture_relationships_with_source<'a, I>(
    slide_xml: &[u8],
    picture_xmls: I,
    max_output_bytes: u64,
) -> Result<Vec<super::PictureRelationship>>
where
    I: IntoIterator<Item = &'a [u8]>,
{
    let layouts = owner::locate_all(slide_xml)?;
    let mut picture_xmls = picture_xmls.into_iter();
    let mut relationships = Vec::new();
    relationships
        .try_reserve_exact(layouts.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed picture relationships",
            source,
        })?;
    for layout in &layouts {
        // The source owner range is authoritative: it comes from the
        // full-slide owner walk, which also validates the picture layout, so
        // relationship parsing uses it rather than the Scene fragment.
        let _scene_picture_xml = picture_xmls
            .next()
            .ok_or_else(|| Error::Invalid("picture layout/source count differs".into()))?;
        let picture_xml = slide_xml
            .get(layout.picture.start..layout.picture.end)
            .ok_or_else(|| Error::Invalid("picture layout range is outside slide XML".into()))?;
        relationships.push(parse_picture_relationship_with_layout(
            layout,
            picture_xml,
            max_output_bytes,
        )?);
    }
    if picture_xmls.next().is_some() {
        return Err(Error::Invalid("picture layout/source count differs".into()));
    }
    Ok(relationships)
}

/// Inventory direct picture relationships from the raw source owner.  This
/// path is used when the general shape Scene cannot preprocess unrelated
/// opaque extension markup (for example a PI or an unsupported MCE branch).
/// The SVG owner scanner still supplies the complete source ranges and keeps
/// the selection semantics identical to the normal path.
pub(super) fn parse_picture_relationships_from_source(
    slide_xml: &[u8],
    max_output_bytes: u64,
) -> Result<Vec<super::PictureRelationship>> {
    let layouts = owner::locate_all(slide_xml)?;
    let mut relationships = Vec::new();
    relationships
        .try_reserve_exact(layouts.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed raw picture relationships",
            source,
        })?;
    for layout in &layouts {
        let picture_xml = slide_xml
            .get(layout.picture.start..layout.picture.end)
            .ok_or_else(|| Error::Invalid("picture layout range is outside slide XML".into()))?;
        relationships.push(parse_picture_relationship_with_layout(
            layout,
            picture_xml,
            max_output_bytes,
        )?);
    }
    Ok(relationships)
}

/// Validate every direct picture owner in a raw slide using the strict
/// recognized-extension grammar.  Commit-time dependency validation uses this
/// single owner walk so a malformed sibling sharing a relationship cannot be
/// hidden by selecting only the edited picture.
pub(super) fn validate_picture_relationships_from_source(
    slide_xml: &[u8],
    max_output_bytes: u64,
) -> Result<()> {
    let layouts = owner::locate_all(slide_xml)?;
    for layout in &layouts {
        let picture_xml = slide_xml
            .get(layout.picture.start..layout.picture.end)
            .ok_or_else(|| Error::Invalid("picture layout range is outside slide XML".into()))?;
        let complete = owner::namespace_complete_picture(layout, picture_xml, max_output_bytes)?;
        super::parse_picture_relationship_strict_with_limit(&complete, max_output_bytes)?;
    }
    Ok(())
}

pub(super) fn namespace_complete_element_fragment(
    fragment: &[u8],
    inherited_namespaces: &[(Vec<u8>, Vec<u8>)],
    max_output_bytes: u64,
) -> Result<Vec<u8>> {
    owner::namespace_complete_element_fragment(fragment, inherited_namespaces, max_output_bytes)
}

fn capture_svg_attachment_snapshot(
    editor: &SourceBackedPresentationEditor,
    slide_position: usize,
    image_position: usize,
    _operation: &'static str,
) -> Result<SourceBackedSvgAttachmentSnapshot> {
    // SVG ownership validates the complete slide with its own bounded raw
    // scanner.  Do not run the legacy transition grammar here: it rejects
    // processing instructions and markup-compatibility payloads nested in
    // opaque extension children that this lifecycle preserves byte-for-byte.
    let slide = editor.publication_raw_slide_snapshot_from_retained_catalog(slide_position)?;
    let view = editor.package.part(&slide.part_uri)?;
    let source_part = super::SourcePart::from_view(&view, view.data()?)?;
    validate_source_slide_root(&source_part)?;
    validate_full_slide_picture_relationships(&editor.package, source_part.blob())?;
    let layout = selected_picture_layout(source_part.blob(), image_position)?;
    let picture_xml = source_part
        .blob()
        .get(layout.picture.start..layout.picture.end)
        .ok_or_else(|| Error::Invalid("selected picture range is outside slide XML".into()))?;
    let relationship_xml =
        owner::namespace_complete_picture(&layout, picture_xml, editor.limits.max_part_bytes())?;
    let relationship = super::parse_picture_relationship_with_limit(
        &relationship_xml,
        editor.limits.max_part_bytes(),
    )?;
    let raster_target = resolve_picture_target(&editor.package, &view, &relationship)?;
    let (raster_part_uri, raster_content_type) = match raster_target {
        super::SourceImageTarget::Internal {
            part_uri,
            content_type,
        } => (part_uri, content_type),
        super::SourceImageTarget::External { .. } => {
            return Err(Error::Relationship(
                "SVG attachment lifecycle requires an internal raster fallback".into(),
            ));
        },
    };
    if !is_png_content_type(&raster_content_type) {
        return Err(Error::ContentType {
            expected: "image/png".into(),
            actual: raster_content_type,
        });
    }
    let raster_relationship = view.rels().get(&relationship.id).ok_or_else(|| {
        Error::Relationship(format!(
            "picture raster relationship '{}' is missing",
            relationship.id
        ))
    })?;
    if raster_relationship.target_mode() != TargetMode::Internal {
        return Err(Error::Relationship(
            "SVG attachment lifecycle requires an internal raster relationship".into(),
        ));
    }
    let svg = if let Some(svg_relationship) = relationship.svg.as_ref() {
        let target = resolve_svg_target(&editor.package, &view, svg_relationship)?;
        let (part_uri, content_type) = match target.target() {
            super::SourceImageTarget::Internal {
                part_uri,
                content_type,
            } => (part_uri.clone(), content_type.to_owned()),
            super::SourceImageTarget::External { .. } => {
                return Err(Error::Relationship(
                    "SVG attachment lifecycle refuses linked SVG targets".into(),
                ));
            },
        };
        if !is_svg_content_type(&content_type) {
            return Err(Error::ContentType {
                expected: "image/svg+xml".into(),
                actual: content_type,
            });
        }
        let media = editor.package.part(&part_uri)?;
        if !media.rels().is_empty() {
            return Err(Error::Relationship(format!(
                "SVG media Part '{}' has outbound relationships",
                part_uri.as_str()
            )));
        }
        let data = media.data()?;
        let relation = view.rels().get(&svg_relationship.id).ok_or_else(|| {
            Error::Relationship(format!(
                "picture SVG relationship '{}' is missing",
                svg_relationship.id
            ))
        })?;
        Some(SourceSvgAttachment {
            relationship_id: svg_relationship.id.clone(),
            relationship_type: relation.reltype().to_owned(),
            part_uri,
            payload: SourcePayload::Original(data),
        })
    } else {
        None
    };
    let (default_relationship_id, default_part_uri) =
        allocate_svg_identity(&editor.package, &view, &raster_part_uri)?;
    let (existing_relationship_ids, physical_member_names) =
        capture_identity_inventory(&editor.package, &view)?;
    Ok(SourceBackedSvgAttachmentSnapshot {
        slide,
        image_position,
        raster_relationship_id: relationship.id,
        raster_relationship_type: raster_relationship.reltype().to_owned(),
        raster_part_uri,
        svg,
        default_relationship_id,
        default_part_uri,
        existing_relationship_ids,
        physical_member_names,
        limits: editor.limits,
    })
}

fn capture_identity_inventory(
    package: &litchi_opc::SourceBackedPackage,
    view: &litchi_opc::PartView<'_>,
) -> Result<(Arc<[String]>, Arc<[String]>)> {
    let mut relationship_ids = Vec::new();
    relationship_ids
        .try_reserve_exact(view.rels().iter().count())
        .map_err(|source| Error::Allocation {
            resource: "SVG attachment relationship identity inventory",
            source,
        })?;
    for relationship in view.rels().iter() {
        relationship_ids.push(super::clone_relationship_text(
            relationship.r_id(),
            "SVG attachment relationship identity",
        )?);
    }
    let mut physical_member_names = Vec::new();
    physical_member_names
        .try_reserve_exact(package.physical_member_names().len())
        .map_err(|source| Error::Allocation {
            resource: "SVG attachment physical identity inventory",
            source,
        })?;
    for name in package.physical_member_names() {
        physical_member_names.push(super::clone_relationship_text(
            name,
            "SVG attachment physical identity",
        )?);
    }
    Ok((
        Arc::from(relationship_ids.into_boxed_slice()),
        Arc::from(physical_member_names.into_boxed_slice()),
    ))
}

fn allocate_svg_identity(
    package: &litchi_opc::SourceBackedPackage,
    view: &litchi_opc::PartView<'_>,
    _raster_uri: &PackURI,
) -> Result<(String, PackURI)> {
    let relationship_id = (0..MAX_GENERATED_NAME_ATTEMPTS)
        .map(|index| {
            if index == 0 {
                "rIdSvg".to_owned()
            } else {
                format!("rIdSvg{index}")
            }
        })
        .find(|candidate| view.rels().get(candidate).is_none())
        .ok_or(Error::Limit {
            resource: "generated SVG relationship ID candidates",
            limit: MAX_GENERATED_NAME_ATTEMPTS,
        })?;
    // Keep the ordinary authoring name stable and recognizable.  Collision
    // probing below still makes this safe when a package already contains a
    // vector.svg member (or a case variant).
    let stem = "vector";
    let mut physical = Vec::new();
    for name in package.physical_member_names() {
        physical
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "SVG attachment physical member index",
                source,
            })?;
        physical.push(name);
    }
    for index in 0..MAX_GENERATED_NAME_ATTEMPTS {
        let suffix = if index == 0 {
            String::new()
        } else {
            format!("{index}")
        };
        let candidate = PackURI::new(format!("/ppt/media/{stem}{suffix}.svg"))
            .map_err(|error| Error::Uri(error.to_string()))?;
        let relationship_uri = candidate
            .rels_uri()
            .map_err(|error| Error::Uri(error.to_string()))?;
        let relationship_member = relationship_uri.membername();
        if !physical
            .iter()
            .any(|name| name.eq_ignore_ascii_case(candidate.membername()))
            && !physical
                .iter()
                .any(|name| name.eq_ignore_ascii_case(relationship_member))
        {
            return Ok((relationship_id, candidate));
        }
    }
    Err(Error::Limit {
        resource: "generated SVG media Part candidates",
        limit: MAX_GENERATED_NAME_ATTEMPTS,
    })
}

fn rewrite_picture_svg(
    slide_xml: &[u8],
    layout: &PictureLayout,
    attachment: Option<&str>,
    limits: &litchi_opc::ReadLimits,
) -> Result<Vec<u8>> {
    // Keep an unprefixed DrawingML owner unprefixed. A synthetic `a:` name
    // would require an inherited declaration that the source may not have;
    // the default namespace already binds the empty prefix.
    let drawing_prefix = layout.blip.prefix.clone();
    let mut replacements = Vec::<(ByteRange, Vec<u8>)>::new();
    replacements
        .try_reserve_exact(1)
        .map_err(|source| Error::Allocation {
            resource: "SVG lifecycle source splice replacements",
            source,
        })?;
    match attachment {
        Some(relationship_id) => {
            if layout.svg_extension.is_some() {
                return Err(Error::Relationship(
                    "selected picture already has an SVG extension".into(),
                ));
            }
            // An existing extLst owns its own namespace context.  Its prefix
            // may differ from the blip prefix (and may even rebind the
            // conventional `a` prefix), so every generated child and close
            // tag inserted into that container must use the selected
            // extLst's resolved DrawingML prefix.
            let extension_prefix = layout
                .ext_list
                .as_ref()
                .map_or(drawing_prefix.as_slice(), |ext_list| {
                    ext_list.prefix.as_slice()
                });
            let ext_len = generated_svg_extension_len(extension_prefix, relationship_id)?;
            let ext_list_len = generated_ext_list_len(&drawing_prefix, ext_len)?;
            let output_len = if let Some(ext_list) = layout.ext_list.as_ref() {
                if ext_list.close_start.is_some() {
                    slide_xml
                        .len()
                        .checked_add(ext_len)
                        .ok_or_else(|| Error::Invalid("SVG slide output size overflows".into()))?
                } else {
                    let old = slide_xml
                        .get(ext_list.range.start..ext_list.range.end)
                        .ok_or_else(|| {
                            Error::Invalid("SVG extLst range is outside slide XML".into())
                        })?;
                    let replacement_len =
                        generated_ext_list_from_empty_len(old, extension_prefix, ext_len)?;
                    slide_xml
                        .len()
                        .checked_sub(old.len())
                        .and_then(|length| length.checked_add(replacement_len))
                        .ok_or_else(|| Error::Invalid("SVG slide output size overflows".into()))?
                }
            } else if layout.blip.close_start.is_some() {
                slide_xml
                    .len()
                    .checked_add(ext_list_len)
                    .ok_or_else(|| Error::Invalid("SVG slide output size overflows".into()))?
            } else {
                let old_len = layout
                    .blip
                    .range
                    .end
                    .checked_sub(layout.blip.range.start)
                    .ok_or_else(|| Error::Invalid("SVG blip range is reversed".into()))?;
                if old_len == 0 {
                    return Err(Error::Invalid("SVG blip range is empty".into()));
                }
                let close_len = 3usize
                    .checked_add(qname_len(&drawing_prefix, b"blip")?)
                    .ok_or_else(|| Error::Invalid("SVG blip close size overflows".into()))?;
                slide_xml
                    .len()
                    .checked_sub(1)
                    .and_then(|length| length.checked_add(ext_list_len))
                    .and_then(|length| length.checked_add(close_len))
                    .ok_or_else(|| Error::Invalid("SVG slide output size overflows".into()))?
            };
            check_slide_output_len(output_len, limits)?;
            let ext = generated_svg_extension(extension_prefix, relationship_id)?;
            if let Some(ext_list) = layout.ext_list.as_ref() {
                if let Some(close_start) = ext_list.close_start {
                    replacements.push((
                        ByteRange {
                            start: close_start,
                            end: close_start,
                        },
                        ext,
                    ));
                } else {
                    let old = slide_xml
                        .get(ext_list.range.start..ext_list.range.end)
                        .ok_or_else(|| {
                            Error::Invalid("extLst range is outside slide XML".into())
                        })?;
                    let replacement = generated_ext_list_from_empty(old, extension_prefix, ext)?;
                    replacements.push((ext_list.range, replacement));
                }
            } else if let Some(close_start) = layout.blip.close_start {
                let ext_list = generated_ext_list(&drawing_prefix, ext)?;
                replacements.push((
                    ByteRange {
                        start: close_start,
                        end: close_start,
                    },
                    ext_list,
                ));
            } else {
                let old = slide_xml
                    .get(layout.blip.range.start..layout.blip.range.end)
                    .ok_or_else(|| Error::Invalid("blip range is outside slide XML".into()))?;
                let replacement = generated_nonempty_blip(old, &drawing_prefix, ext)?;
                replacements.push((layout.blip.range, replacement));
            }
        },
        None => {
            let svg = layout.svg_extension.ok_or_else(|| {
                Error::Relationship("selected picture has no SVG extension".into())
            })?;
            if let Some(ext_list) = layout.ext_list.as_ref() {
                let only_svg = ext_list_contains_only_svg(slide_xml, ext_list, svg)?;
                if only_svg {
                    replacements.push((ext_list.range, Vec::new()));
                } else {
                    replacements.push((svg, Vec::new()));
                }
            } else {
                return Err(Error::Relationship(
                    "selected picture SVG extension has no extLst owner".into(),
                ));
            }
        },
    }
    splice_checked(slide_xml, replacements, limits)
}

fn generated_svg_extension(drawing_prefix: &[u8], relationship_id: &str) -> Result<Vec<u8>> {
    validate_relationship_id(relationship_id)?;
    let capacity = generated_svg_extension_len(drawing_prefix, relationship_id)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "generated SVG extension",
            source,
        })?;
    push_open_qname(&mut output, drawing_prefix, b"ext");
    output.extend_from_slice(b" uri=\"{96DAC541-7B7A-43D3-8B79-37D633B846F1}\"><asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" xmlns:r=\"");
    output.extend_from_slice(RELATIONSHIP_NAMESPACE);
    output.extend_from_slice(b"\" r:embed=\"");
    output.extend_from_slice(relationship_id.as_bytes());
    output.extend_from_slice(b"\"/></");
    push_qname(&mut output, drawing_prefix, b"ext");
    output.push(b'>');
    Ok(output)
}

fn qname_len(prefix: &[u8], local: &[u8]) -> Result<usize> {
    prefix
        .len()
        .checked_add(local.len())
        .and_then(|length| length.checked_add(usize::from(!prefix.is_empty())))
        .ok_or_else(|| Error::Invalid("generated XML qualified name size overflows".into()))
}

fn generated_svg_extension_len(drawing_prefix: &[u8], relationship_id: &str) -> Result<usize> {
    validate_relationship_id(relationship_id)?;
    let ext_qname_len = qname_len(drawing_prefix, b"ext")?;
    1usize
        .checked_add(ext_qname_len)
        .and_then(|length| length.checked_add(b" uri=\"{96DAC541-7B7A-43D3-8B79-37D633B846F1}\"><asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" xmlns:r=\"".len()))
        .and_then(|length| length.checked_add(RELATIONSHIP_NAMESPACE.len()))
        .and_then(|length| length.checked_add(b"\" r:embed=\"".len()))
        .and_then(|length| length.checked_add(relationship_id.len()))
        .and_then(|length| length.checked_add(b"\"/></".len()))
        .and_then(|length| length.checked_add(ext_qname_len))
        .and_then(|length| length.checked_add(1))
        .ok_or_else(|| Error::Invalid("generated SVG extension size overflows".into()))
}

fn generated_ext_list_len(drawing_prefix: &[u8], ext_len: usize) -> Result<usize> {
    qname_len(drawing_prefix, b"extLst")?
        .checked_mul(2)
        .and_then(|length| length.checked_add(ext_len))
        .and_then(|length| length.checked_add(5))
        .ok_or_else(|| Error::Invalid("generated extLst size overflows".into()))
}

fn check_slide_output_len(output_len: usize, limits: &litchi_opc::ReadLimits) -> Result<()> {
    if output_len as u64 > limits.max_part_bytes() {
        return Err(Error::Limit {
            resource: "SVG lifecycle slide XML bytes",
            limit: usize::try_from(limits.max_part_bytes()).unwrap_or(usize::MAX),
        });
    }
    Ok(())
}

fn generated_ext_list(drawing_prefix: &[u8], ext: Vec<u8>) -> Result<Vec<u8>> {
    let capacity = generated_ext_list_len(drawing_prefix, ext.len())?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "generated SVG extension list",
            source,
        })?;
    push_open_qname(&mut output, drawing_prefix, b"extLst");
    output.push(b'>');
    output.extend_from_slice(&ext);
    output.extend_from_slice(b"</");
    push_qname(&mut output, drawing_prefix, b"extLst");
    output.push(b'>');
    Ok(output)
}

fn generated_ext_list_from_empty(
    old: &[u8],
    drawing_prefix: &[u8],
    ext: Vec<u8>,
) -> Result<Vec<u8>> {
    let slash = old
        .iter()
        .rposition(|byte| *byte == b'/')
        .ok_or_else(|| Error::Invalid("self-closing extLst has no close slash".into()))?;
    let capacity = generated_ext_list_from_empty_len(old, drawing_prefix, ext.len())?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "generated SVG extension list",
            source,
        })?;
    output.extend_from_slice(&old[..slash]);
    output.push(b'>');
    output.extend_from_slice(&ext);
    output.extend_from_slice(b"</");
    push_qname(&mut output, drawing_prefix, b"extLst");
    output.push(b'>');
    Ok(output)
}

fn generated_ext_list_from_empty_len(
    old: &[u8],
    drawing_prefix: &[u8],
    ext_len: usize,
) -> Result<usize> {
    if !old.ends_with(b"/>") {
        return Err(Error::Invalid(
            "empty extLst is not self-closing in source".into(),
        ));
    }
    let close_len = 3usize
        .checked_add(qname_len(drawing_prefix, b"extLst")?)
        .ok_or_else(|| Error::Invalid("generated extLst close size overflows".into()))?;
    old.len()
        .checked_sub(1)
        .and_then(|length| length.checked_add(ext_len))
        .and_then(|length| length.checked_add(close_len))
        .ok_or_else(|| Error::Invalid("generated extLst size overflows".into()))
}

fn generated_nonempty_blip(old: &[u8], drawing_prefix: &[u8], ext: Vec<u8>) -> Result<Vec<u8>> {
    if !old.ends_with(b"/>") {
        return Err(Error::Invalid(
            "empty blip is not self-closing in source".into(),
        ));
    }
    let slash = old
        .iter()
        .rposition(|byte| *byte == b'/')
        .ok_or_else(|| Error::Invalid("self-closing blip has no close slash".into()))?;
    let ext_list = generated_ext_list(drawing_prefix, ext)?;
    let close_len = 3usize
        .checked_add(qname_len(drawing_prefix, b"blip")?)
        .ok_or_else(|| Error::Invalid("generated nonempty blip close size overflows".into()))?;
    let capacity = old
        .len()
        .checked_sub(1)
        .and_then(|value| value.checked_add(ext_list.len()))
        .and_then(|value| value.checked_add(close_len))
        .ok_or_else(|| Error::Invalid("generated nonempty blip size overflows".into()))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "generated nonempty SVG blip",
            source,
        })?;
    output.extend_from_slice(&old[..slash]);
    output.push(b'>');
    output.extend_from_slice(&ext_list);
    output.extend_from_slice(b"</");
    push_qname(&mut output, drawing_prefix, b"blip");
    output.push(b'>');
    Ok(output)
}

fn push_open_qname(output: &mut Vec<u8>, prefix: &[u8], local: &[u8]) {
    output.push(b'<');
    push_qname(output, prefix, local);
}

fn push_qname(output: &mut Vec<u8>, prefix: &[u8], local: &[u8]) {
    if !prefix.is_empty() {
        output.extend_from_slice(prefix);
        output.push(b':');
    }
    output.extend_from_slice(local);
}

fn ext_list_contains_only_svg(xml: &[u8], ext_list: &ElementRange, svg: ByteRange) -> Result<bool> {
    // Keep the wrapper whenever its opening tag carries any attribute,
    // including namespace declarations.  The declaration bytes are part of
    // the source-owned wrapper and dropping them would make an inverse
    // detach lose lexical metadata (and can change the namespace context for
    // unrelated opaque payload).
    let opening = xml
        .get(ext_list.range.start..ext_list.start_end)
        .ok_or_else(|| Error::Invalid("extLst opening range is outside slide XML".into()))?;
    let mut opening_reader = NsReader::from_reader(opening);
    let opening_event = opening_reader
        .read_event()
        .map_err(|error| Error::Xml(error.to_string()))?;
    if let Event::Start(element) | Event::Empty(element) = opening_event {
        if element.checked_attributes().next().is_some() {
            return Ok(false);
        }
    } else {
        return Err(Error::Invalid(
            "extLst opening range is not an element".into(),
        ));
    }
    let end = ext_list.close_start.unwrap_or(ext_list.range.end);
    let children = xml
        .get(ext_list.start_end..end)
        .ok_or_else(|| Error::Invalid("extLst child range is outside slide XML".into()))?;
    let mut reader = NsReader::from_reader(children);
    let origin = ReaderOrigin::of(children);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut saw_other = false;
    let mut depth = 0usize;
    let mut direct_ext_start = None;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| Error::Invalid("extLst child offset exceeds usize".into()))?;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| Error::Invalid("extLst child end exceeds usize".into()))?;
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if element.local_name().as_ref() == b"ext" {
                        direct_ext_start = Some(start);
                    } else {
                        saw_other = true;
                    }
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("SVG extLst child depth overflows".into()))?;
            },
            Event::Empty(element) => {
                if depth == 0 {
                    if element.local_name().as_ref() == b"ext" {
                        let absolute = ByteRange {
                            start: ext_list.start_end + start,
                            end: ext_list.start_end + end,
                        };
                        if absolute.start != svg.start || absolute.end != svg.end {
                            saw_other = true;
                        }
                    } else {
                        saw_other = true;
                    }
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Invalid("SVG extLst child end is unmatched".into()))?;
                if depth == 0 {
                    if let Some(child_start) = direct_ext_start.take() {
                        let absolute = ByteRange {
                            start: ext_list.start_end + child_start,
                            end: ext_list.start_end + end,
                        };
                        if absolute.start != svg.start || absolute.end != svg.end {
                            saw_other = true;
                        }
                    }
                }
            },
            Event::Text(text) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                saw_other = true;
            },
            // A comment is opaque payload in the recognized container. Keep
            // the wrapper when it is the only sibling beside the SVG so that
            // detach never discards unrelated source bytes.
            Event::Comment(_) | Event::CData(_) | Event::GeneralRef(_) | Event::PI(_) => {
                saw_other = true
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(!saw_other)
}

fn splice_checked(
    source: &[u8],
    mut replacements: Vec<(ByteRange, Vec<u8>)>,
    limits: &litchi_opc::ReadLimits,
) -> Result<Vec<u8>> {
    replacements.sort_unstable_by_key(|(range, _)| range.start);
    let mut output_len = source.len();
    let mut previous_end = 0usize;
    for (range, replacement) in &replacements {
        if range.start < previous_end || range.end < range.start || range.end > source.len() {
            return Err(Error::Invalid(
                "SVG source splice ranges overlap or exceed source".into(),
            ));
        }
        output_len = output_len
            .checked_sub(range.end - range.start)
            .and_then(|value| value.checked_add(replacement.len()))
            .ok_or_else(|| Error::Invalid("SVG source splice output size overflows".into()))?;
        previous_end = range.end;
    }
    if output_len as u64 > limits.max_part_bytes() {
        return Err(Error::Limit {
            resource: "SVG lifecycle slide XML bytes",
            limit: usize::try_from(limits.max_part_bytes()).unwrap_or(usize::MAX),
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "SVG lifecycle slide XML",
            source,
        })?;
    let mut cursor = 0usize;
    for (range, replacement) in replacements {
        output.extend_from_slice(&source[cursor..range.start]);
        output.extend_from_slice(&replacement);
        cursor = range.end;
    }
    output.extend_from_slice(&source[cursor..]);
    Ok(output)
}

fn build_attachment_topology_plan(
    package: &litchi_opc::SourceBackedPackage,
    current: &SourceBackedSvgAttachmentSnapshot,
    target: &SourceBackedSvgAttachmentSnapshot,
) -> Result<SourceTopologyPlan> {
    if current.slide.position() != target.slide.position()
        || current.image_position != target.image_position
        || current.raster_relationship_id != target.raster_relationship_id
        || current.raster_relationship_type != target.raster_relationship_type
        || current.raster_part_uri != target.raster_part_uri
    {
        return Err(Error::StaleSource);
    }
    let owner = current.slide.part_uri.clone();
    let mut plan = SourceTopologyPlan::new();
    let slide_bytes =
        target.slide.xml.edited_bytes().ok_or_else(|| {
            Error::Invalid("SVG lifecycle target has no edited slide payload".into())
        })?;
    // Revalidate every pre-edit owner before accepting a commit.  The
    // capture path may select the first valid owner while retaining unrelated
    // opaque bytes; a malformed sibling sharing a relationship must still
    // fail closed before the commit becomes publishable.
    validate_picture_relationships_from_source(
        current.slide.xml.as_bytes(),
        current.limits.max_part_bytes(),
    )?;
    validate_full_slide_picture_relationships(package, slide_bytes.as_slice())?;
    let target_layout = selected_picture_layout(slide_bytes.as_slice(), target.image_position)?;
    let target_picture = slide_bytes
        .get(target_layout.picture.start..target_layout.picture.end)
        .ok_or_else(|| {
            Error::Invalid("SVG lifecycle target picture range is outside slide".into())
        })?;
    let target_picture_xml = owner::namespace_complete_picture(
        &target_layout,
        target_picture,
        current.limits.max_part_bytes(),
    )?;
    super::parse_picture_relationship_strict_with_limit(
        &target_picture_xml,
        current.limits.max_part_bytes(),
    )?;
    let source_xml = package.part(&owner)?.source_xml()?;
    let source_root = root_element_range(source_xml.bytes())?;
    let target_root = root_element_range(slide_bytes.as_slice())?;
    let source_root_bytes = source_xml
        .bytes()
        .get(source_root.start..source_root.end)
        .ok_or_else(|| Error::Invalid("source slide root range is outside source".into()))?;
    let target_root_bytes = slide_bytes
        .as_slice()
        .get(target_root.start..target_root.end)
        .ok_or_else(|| Error::Invalid("target slide root range is outside source".into()))?;
    let proof = source_xml.checked_range(source_root.start..source_root.end, source_root_bytes)?;
    let fragment = AuthoredXmlFragment::markup(target_root_bytes.to_vec())?;
    let mut publication = source_xml.into_publication()?;
    publication.replace(proof, fragment)?;
    plan.try_replace_source_xml_part(owner.clone(), publication.finish()?)?;
    match (&current.svg, &target.svg) {
        (None, Some(svg)) => {
            validate_svg_payload(&current.limits, svg.bytes().len())?;
            let existing = package.part(&svg.part_uri);
            if existing.is_ok() {
                return Err(Error::Relationship(format!(
                    "SVG attachment target Part '{}' already exists",
                    svg.part_uri.as_str()
                )));
            }
            plan.try_add_part_shared(
                svg.part_uri.clone(),
                "image/svg+xml".to_owned(),
                match &svg.payload {
                    SourcePayload::Edited(bytes) => Arc::clone(bytes),
                    SourcePayload::Original(_) => {
                        return Err(Error::Invalid(
                            "new SVG attachment payload is not edit-owned".into(),
                        ));
                    },
                },
            )?;
            plan.try_add_internal_relationship(
                owner,
                svg.relationship_id.clone(),
                svg.relationship_type.clone(),
                svg.part_uri.clone(),
            )?;
        },
        (Some(svg), None) => {
            let relation_id = svg.relationship_id.clone();
            if !relationship_id_is_referenced_elsewhere(
                current.slide.xml.as_bytes(),
                &relation_id,
                current.image_position,
            )? {
                plan.try_remove_relationship(owner.clone(), relation_id)?;
                super::remove_svg_media_if_unreferenced(
                    package,
                    &mut plan,
                    &owner,
                    &svg.relationship_id,
                    &svg.part_uri,
                )?;
            }
        },
        (Some(_), Some(_)) => {
            return Err(Error::UnsafeEdit {
                operation: "SVG lifecycle publication",
                reason: "attach/detach transactions cannot replace an existing SVG; use edit_svg_image",
            });
        },
        (None, None) => {},
    }
    Ok(plan)
}

fn root_element_range(xml: &[u8]) -> Result<ByteRange> {
    // quick-xml reports offsets relative to the byte stream it consumes.  A
    // UTF-8 BOM is legal before the document element but is not represented
    // by an event, so source ranges must include it when they are applied to
    // the original slide bytes: the reader origin adds it.
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root = None;
    let mut buffer = Vec::new();
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| Error::Invalid("slide XML root offset exceeds usize".into()))?;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| Error::Invalid("slide XML root end exceeds usize".into()))?;
        match event {
            Event::Start(_) => {
                if depth == 0 {
                    if root.is_some() {
                        return Err(Error::Invalid(
                            "slide XML contains more than one root element".into(),
                        ));
                    }
                    root = Some(start);
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("slide XML root depth overflows".into()))?;
            },
            Event::Empty(_) if depth == 0 => {
                if root.is_some() {
                    return Err(Error::Invalid(
                        "slide XML contains more than one root element".into(),
                    ));
                }
                return Ok(ByteRange { start, end });
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Invalid("slide XML has an unmatched end".into()))?;
                if depth == 0 {
                    let start = root
                        .take()
                        .ok_or_else(|| Error::Invalid("slide XML has no root element".into()))?;
                    return Ok(ByteRange { start, end });
                }
            },
            Event::Eof => break,
            Event::Decl(_) | Event::DocType(_) | Event::PI(_) => {},
            _ => {},
        }
    }
    Err(Error::Invalid("slide XML has no root element".into()))
}

fn relationship_id_is_referenced_elsewhere(
    slide_xml: &[u8],
    relationship_id: &str,
    selected_image_position: usize,
) -> Result<bool> {
    let selected = selected_picture_layout(slide_xml, selected_image_position)?;
    let selected_extension = selected.svg_extension;
    let mut reader = Reader::from_reader(slide_xml);
    let origin = ReaderOrigin::of(slide_xml);
    let mut buffer = Vec::new();
    loop {
        let event_start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| Error::Invalid("SVG relationship scan position overflows".into()))?;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Xml(error.to_string()))?;
        let event_end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| Error::Invalid("SVG relationship scan position overflows".into()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                // The selected native SVG extension necessarily contains the
                // relationship ID in its own r:embed attribute.  Ignore that
                // exact source range, while retaining every other decoded
                // attribute reference (including opaque foreign attributes).
                let in_selected_extension = selected_extension
                    .is_some_and(|range| event_start >= range.start && event_end <= range.end);
                if !in_selected_extension {
                    for attribute in element.checked_attributes() {
                        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        // Unknown relationship-bearing attributes are opaque
                        // to this lifecycle.  Retain the relationship for
                        // any decoded occurrence, including character
                        // references and whitespace-delimited ID lists,
                        // rather than risk leaving a dangling reference.
                        if value.as_ref().contains(relationship_id) {
                            return Ok(true);
                        }
                    }
                }
            },
            Event::Eof => return Ok(false),
            _ => {},
        }
        buffer.clear();
    }
}
