//! Detached snapshots for slide related parts.
//!
//! These values are detached from the source owner.  Acquisition parses and
//! validates the selected part eagerly, so the snapshots remain usable after
//! the source package has been converted into an editor package.
//! Cloning a part snapshot shares its retained data. Original source bytes keep
//! their managed memory reservation until the last retaining handle is dropped.
//! Native package identities and bytes require explicit [`raw`] access.
//!
//! ```
//! use std::io::Cursor;
//! use litchi_pptx::SourceBackedPresentation;
//!
//! let mut builder = litchi_pptx::builder::Builder::new()?;
//! builder.presentation()?.add_slide()?.set_notes("Speaker notes");
//! let source = SourceBackedPresentation::from_reader(Cursor::new(builder.build()?.to_bytes()?))?;
//! let notes = source.slide(0).unwrap().notes()?.unwrap();
//! drop(source);
//! assert_eq!(notes.text(), Some("Speaker notes"));
//! # Ok::<(), litchi_pptx::Error>(())
//! ```

use std::borrow::Cow;
use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_opc::SourceLimits;
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcError, PackURI, Part, PartData, SourceLimitKind, TargetMode};

use crate::chart::Type as ChartType;
use crate::comments::{
    AUTHORS_CONTENT_TYPE, AUTHORS_REL, COMMENTS_REL, parse_comment_authors,
    parse_slide_comments_strict,
};
use crate::error::{Error, Result};
use crate::master_layout::SlideLayoutKind;
use crate::parts::{SlideLayoutPart, SlidePart, is_relationship_type};
use crate::presentation::source::{SourceInner, SourcePart, SourceSlide};
use crate::shape::owned::Scene as OwnedScene;
use litchi_opc::Relationship;

mod catalog;
mod comments;
pub(crate) mod eager_notes;
mod notes;
pub use catalog::Master;
pub use comments::{Comment, Comments};
pub mod raw;

const NOTES_SLIDE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const SLIDE_LAYOUT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const CHART_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

const CT_SLIDE: &str = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const CT_SLIDE_LAYOUT: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
const CT_NOTES_SLIDE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
const CT_COMMENTS: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
const CT_CHART: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const MAX_CHARTS_PER_SLIDE: usize = 4096;
const MAX_CHART_XML: usize = 64 * 1024 * 1024;

pub(crate) fn load_eager_notes<'a>(
    package: &'a litchi_opc::OpcPackage,
    slide: &SlidePart<'a>,
    shape_limits: crate::shape::Limits,
    notes_limits: crate::notes::Limits,
) -> Result<Option<Notes>> {
    eager_notes::load(package, slide, shape_limits, notes_limits)
}

pub(crate) fn load_eager_comments<'a>(
    package: &'a litchi_opc::OpcPackage,
    slide: &SlidePart<'a>,
    limits: crate::comments::Limits,
) -> Result<Option<Comments>> {
    comments::load_eager(package, slide, limits)
}

/// A source part retained by a detached related-part snapshot.
#[derive(Clone)]
pub(crate) struct RelatedPart {
    pub(crate) partname: PackURI,
    pub(crate) relationship_id: Box<str>,
    pub(crate) data: RelatedData,
}

/// Retained bytes for a source-backed or eager related-part snapshot.
#[derive(Clone)]
pub(crate) enum RelatedData {
    Source(PartData),
    Eager(Arc<Vec<u8>>),
}

impl RelatedData {
    pub(crate) fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Source(data) => data.as_bytes(),
            Self::Eager(data) => data.as_slice(),
        }
    }
}

impl From<PartData> for RelatedData {
    fn from(data: PartData) -> Self {
        Self::Source(data)
    }
}

impl RelatedPart {
    pub(super) fn new(part: SourcePart, relationship_id: &str) -> Self {
        Self {
            partname: part.source_partname().clone(),
            relationship_id: relationship_id.into(),
            data: part.source_data().clone().into(),
        }
    }

    pub(crate) fn from_eager(part: &dyn Part, relationship_id: &str) -> Self {
        Self {
            partname: part.partname().clone(),
            relationship_id: relationship_id.into(),
            data: RelatedData::Eager(part.blob_arc()),
        }
    }
}

impl RelatedPart {
    pub(crate) fn as_bytes(&self) -> &[u8] {
        self.data.as_bytes()
    }
}

/// A detached source-backed slide layout.
#[derive(Clone)]
pub struct Layout {
    data: Arc<LayoutData>,
}

struct LayoutData {
    related: RelatedPart,
    name: Option<String>,
    kind: Option<SlideLayoutKind>,
    scene: OwnedScene,
}

impl fmt::Debug for Layout {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Layout").finish_non_exhaustive()
    }
}

impl Layout {
    /// The authored layout name, if present.
    pub fn name(&self) -> Option<&str> {
        self.data.name.as_deref()
    }

    /// The typed layout kind, if the authored token is recognized.
    pub fn kind(&self) -> Option<SlideLayoutKind> {
        self.data.kind
    }

    /// The detached shape scene retained by this snapshot.
    pub fn scene(&self) -> &OwnedScene {
        &self.data.scene
    }
}

/// A detached notes-slide snapshot shared by eager and source-backed reads.
#[derive(Clone)]
pub struct Notes {
    data: Arc<NotesData>,
}

struct NotesData {
    related: RelatedPart,
    text: Option<String>,
    scene: OwnedScene,
}

impl fmt::Debug for Notes {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Notes").finish_non_exhaustive()
    }
}

impl Notes {
    pub(crate) fn from_parts(
        related: RelatedPart,
        text: Option<String>,
        scene: OwnedScene,
    ) -> Self {
        Self {
            data: Arc::new(NotesData {
                related,
                text,
                scene,
            }),
        }
    }

    /// The primary body placeholder's text, if present.
    pub fn text(&self) -> Option<&str> {
        self.data.text.as_deref()
    }

    /// The detached shape scene retained by this snapshot.
    pub fn scene(&self) -> &OwnedScene {
        &self.data.scene
    }
}

/// A detached source-backed chart summary.
#[derive(Clone)]
pub struct Chart {
    data: Arc<ChartData>,
}

struct ChartData {
    related: RelatedPart,
    kind: ChartType,
    title: Option<String>,
    has_legend: bool,
}

impl fmt::Debug for Chart {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Chart").finish_non_exhaustive()
    }
}

impl Chart {
    /// The chart's typed kind.
    pub fn kind(&self) -> ChartType {
        self.data.kind
    }

    /// The chart title, if present.
    pub fn title(&self) -> Option<&str> {
        self.data.title.as_deref()
    }

    /// Whether the chart has a legend.
    pub fn has_legend(&self) -> bool {
        self.data.has_legend
    }
}

/// Load a selected source part after enforcing package and XML budgets.
pub(super) fn load_part(
    owner: &SourceInner,
    uri: &PackURI,
    content_type: &str,
    semantic_maximum: usize,
) -> Result<SourcePart> {
    owner.package.check_execution()?;
    let view = owner.package.part(uri)?;
    if view.content_type() != content_type {
        return Err(Error::ContentType {
            expected: content_type.to_owned(),
            actual: view.content_type().to_owned(),
        });
    }

    let declared = view.declared_uncompressed_size()?;
    let physical_limit = owner.limits.max_part_bytes();
    if declared > physical_limit {
        return Err(Error::Opc(OpcError::Limit {
            kind: SourceLimitKind::PartBytes,
            observed: declared,
            maximum: physical_limit,
        }));
    }
    if declared > semantic_maximum as u64 {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::PartBytes),
            observed: usize::try_from(declared).unwrap_or(usize::MAX),
            maximum: semantic_maximum,
        });
    }

    let data = view.data()?;
    crate::source_limits::validate_xml_limits(data.as_bytes(), owner.limits)?;
    // Namespace promotion during MCE can expand XML. Bound that allocation
    // before invoking leaf readers, which otherwise use wider default limits.
    let maximum = semantic_maximum.min(usize::try_from(physical_limit).unwrap_or(usize::MAX));
    let limits = litchi_ooxml_common::mce::Limits {
        max_input_bytes: maximum,
        max_output_bytes: maximum,
        ..Default::default()
    };
    let processed = litchi_ooxml_common::mce::process_markup_compatibility(
        data.as_bytes(),
        &litchi_ooxml_common::mce::Capabilities::default(),
        &limits,
    )?;
    let replacement = match processed.xml {
        Cow::Borrowed(_) => None,
        Cow::Owned(xml) => Some(xml),
    };
    owner.package.check_execution()?;
    let mut part = SourcePart::from_view(&view, data)?;
    part.replacement = replacement;
    Ok(part)
}

/// Process one eager related part with the same XML and bounded-MCE policy
/// used by [`load_part`]. The returned blob is suitable for a leaf parser;
/// callers retain the original package part separately when they need raw
/// related-part bytes for the detached snapshot.
pub(super) fn process_eager_part(
    part: &dyn Part,
    maximum: usize,
    source_limits: SourceLimits,
) -> Result<BlobPart> {
    if part.blob().len() > maximum {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::PartBytes),
            observed: part.blob().len(),
            maximum,
        });
    }
    crate::source_limits::validate_xml_limits(part.blob(), source_limits)?;
    let limits = litchi_ooxml_common::mce::Limits {
        max_input_bytes: maximum,
        max_output_bytes: maximum,
        ..Default::default()
    };
    let processed = litchi_ooxml_common::mce::process_markup_compatibility(
        part.blob(),
        &litchi_ooxml_common::mce::Capabilities::default(),
        &limits,
    )?;
    let blob = match processed.xml {
        Cow::Borrowed(_) => part.blob_arc(),
        Cow::Owned(xml) => Arc::new(xml),
    };
    let mut owned = BlobPart::new_shared(
        part.partname().clone(),
        part.content_type().to_owned(),
        blob,
    );
    for relationship in part.rels().iter() {
        owned.rels_mut().try_add_relationship(
            relationship.reltype().to_owned(),
            relationship.target_ref().to_owned(),
            relationship.r_id().to_owned(),
            relationship.target_mode(),
        )?;
    }
    Ok(owned)
}

fn fence(owner: &SourceInner) -> Result<()> {
    owner.package.check_execution()?;
    owner.package.source_version()?;
    Ok(())
}

pub(super) fn with_fences<T>(
    owner: &SourceInner,
    operation: impl FnOnce() -> Result<T>,
) -> Result<T> {
    fence(owner)?;
    let value = operation();
    fence(owner)?;
    value
}

fn strict_scene(scene: &OwnedScene, resource: &'static str) -> Result<()> {
    if scene.tolerance_report().is_clean() {
        Ok(())
    } else {
        Err(Error::Invalid(match resource {
            "layout" => "source layout scene requires strict shape XML".into(),
            "master" => "source master scene requires strict shape XML".into(),
            _ => "source notes scene requires strict shape XML".into(),
        }))
    }
}

struct SceneCandidate {
    key: PackURI,
    scene: OwnedScene,
    needs_publish: bool,
}

fn acquire_scene(
    owner: &SourceInner,
    part: &SourcePart,
    parse: impl FnOnce(&SourcePart) -> Result<OwnedScene>,
) -> Result<SceneCandidate> {
    let key = part.source_partname().clone();
    if let Some(scene) = owner.scenes.lookup(&key) {
        return Ok(SceneCandidate {
            key,
            scene,
            needs_publish: false,
        });
    }

    let scene = parse(part)?;
    Ok(SceneCandidate {
        key,
        scene,
        needs_publish: true,
    })
}

fn publish_scene(owner: &SourceInner, candidate: SceneCandidate) -> Result<OwnedScene> {
    let SceneCandidate {
        key,
        scene,
        needs_publish,
    } = candidate;
    let scene = if needs_publish {
        owner.scenes.publish(&key, scene)
    } else {
        scene
    };
    if let Err(error) = fence(owner) {
        if needs_publish {
            owner.scenes.discard(&key, &scene);
        }
        return Err(error);
    }
    Ok(scene)
}

fn selected_relationship(
    owner: &SourceInner,
    slide: &SourceSlide,
    relationship_type: &str,
    local_type: &str,
    content_type: &str,
    semantic_maximum: usize,
    duplicate_message: &'static str,
) -> Result<Option<(String, SourcePart)>> {
    let slide_part = owner.package.part(&slide.data.part_uri)?;
    let mut selected: Option<&Relationship> = None;
    for relationship in slide_part.rels().iter() {
        if !is_relationship_type(relationship.reltype(), relationship_type, local_type) {
            continue;
        }
        if selected.is_some() {
            return Err(Error::Invalid(duplicate_message.into()));
        }
        selected = Some(relationship);
    }

    let Some(relationship) = selected else {
        return Ok(None);
    };
    if relationship.target_mode() == TargetMode::External {
        return Err(Error::Invalid(
            "source related relationship must be internal".into(),
        ));
    }
    let target = relationship.target_partname()?;
    let part = load_part(owner, &target, content_type, semantic_maximum)?;
    Ok(Some((relationship.r_id().to_owned(), part)))
}

fn selected_source_authors(
    owner: &SourceInner,
    comments: &[crate::comments::Comment],
    limits: crate::comments::Limits,
) -> Result<Option<(String, SourcePart, Vec<crate::comments::Author>)>> {
    // A non-empty legacy comments part requires the presentation-level author
    // relationship and its selected closure before any values are projected.
    if comments.is_empty() {
        return Ok(None);
    }

    let mut selected = None;
    for relationship in owner._presentation.source_relationships().iter() {
        owner.package.check_execution()?;
        if !is_relationship_type(relationship.reltype(), AUTHORS_REL, "commentAuthors") {
            continue;
        }
        if selected.is_some() {
            return Err(Error::Relationship(
                "source presentation has multiple comment-author relationships".into(),
            ));
        }
        if relationship.target_mode() == TargetMode::External {
            return Err(Error::Relationship(
                "source comment-author relationship must be internal".into(),
            ));
        }
        selected = Some(relationship);
    }

    let relationship = selected.ok_or_else(|| {
        Error::Relationship("source comments require a comment-author part".into())
    })?;
    let target = relationship.target_partname().map_err(|_| {
        Error::Relationship("source comment-author relationship target is invalid".into())
    })?;
    let part = load_part(owner, &target, AUTHORS_CONTENT_TYPE, limits.max_input_bytes)?;
    if !part.source_relationships().is_empty() {
        return Err(Error::Relationship(
            "source comment-author part must not have relationships".into(),
        ));
    }
    let authors = parse_comment_authors(part.blob(), limits)?;
    Ok(Some((relationship.r_id().to_owned(), part, authors)))
}

fn chart_relationships(rels: &litchi_opc::Relationships) -> Result<Vec<&Relationship>> {
    let selected = || {
        rels.iter()
            .filter(|relationship| is_relationship_type(relationship.reltype(), CHART_REL, "chart"))
    };
    let count = selected().count();
    let limit = |observed| Error::Limit {
        kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Relationships),
        observed,
        maximum: MAX_CHARTS_PER_SLIDE,
    };
    if count > MAX_CHARTS_PER_SLIDE {
        return Err(limit(count));
    }
    let mut relationships = Vec::new();
    relationships
        .try_reserve_exact(count)
        .map_err(|source| Error::Allocation {
            resource: "source chart relationship snapshots",
            source,
        })?;
    relationships.extend(selected());
    relationships.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
    Ok(relationships)
}

fn read_chart(owner: &SourceInner, relationship: &Relationship) -> Result<Chart> {
    if relationship.target_mode() == TargetMode::External {
        return Err(Error::Invalid(
            "source chart relationship must be internal".into(),
        ));
    }
    let target = relationship.target_partname()?;
    let part = load_part(owner, &target, CT_CHART, MAX_CHART_XML)?;
    let info = crate::chart::Part::from_part(&part)?.chart_info()?;
    if !info.tolerance_report().is_clean() || info.chart_type == ChartType::Unknown {
        return Err(Error::Invalid(
            "source chart requires strict chart XML".into(),
        ));
    }
    Ok(Chart {
        data: Arc::new(ChartData {
            related: RelatedPart::new(part, relationship.r_id()),
            kind: info.chart_type,
            title: info.title,
            has_legend: info.has_legend,
        }),
    })
}

pub(crate) fn notes_text(scene: &OwnedScene) -> Result<Option<String>> {
    let mut seen_body = false;
    let mut text = None;
    for handle in scene.roots() {
        let shape = handle.view();
        if !matches!(shape.placeholder(), Some(placeholder) if placeholder.kind() == Some("body")) {
            continue;
        }
        if seen_body {
            return Err(Error::Invalid(
                "notes slide has multiple body placeholders".into(),
            ));
        }
        seen_body = true;
        if let Some(body) = shape.text_body()? {
            let value = body.text();
            if !value.is_empty() {
                text = Some(value);
            }
        }
    }
    Ok(text)
}

struct LayoutPending {
    related: RelatedPart,
    name: Option<String>,
    kind: Option<SlideLayoutKind>,
    candidate: SceneCandidate,
}

impl LayoutPending {
    fn into_layout(self) -> Layout {
        Layout {
            data: Arc::new(LayoutData {
                related: self.related,
                name: self.name,
                kind: self.kind,
                scene: self.candidate.scene,
            }),
        }
    }
}

fn prepare_layout(
    owner: &SourceInner,
    part: SourcePart,
    relationship_id: &str,
) -> Result<LayoutPending> {
    let layout = SlideLayoutPart::from_part(&part)?;
    let name = layout.authored_name()?;
    let kind = match layout.kind()? {
        None => None,
        Some(token) => Some(
            SlideLayoutKind::from_token(&token)
                .ok_or_else(|| Error::Invalid("source layout kind is unsupported".into()))?,
        ),
    };
    let candidate = acquire_scene(owner, &part, |source| {
        let layout = SlideLayoutPart::from_part(source)?;
        let scene = OwnedScene::from_source(
            layout.shapes_with_limits(
                crate::open_options::selected(owner.semantic_limits.as_ref()).shape(),
            )?,
            source.source_data().clone(),
        )?;
        strict_scene(&scene, "layout")?;
        Ok(scene)
    })?;
    Ok(LayoutPending {
        related: RelatedPart::new(part, relationship_id),
        name,
        kind,
        candidate,
    })
}

fn finish_layout(owner: &SourceInner, pending: LayoutPending) -> Result<Layout> {
    let scene = publish_scene(owner, pending.candidate)?;
    Ok(Layout {
        data: Arc::new(LayoutData {
            related: pending.related,
            name: pending.name,
            kind: pending.kind,
            scene,
        }),
    })
}

/// Acquire one layout Part selected by an effective slide binding.
///
/// Structural primary revisions retain the source layout dependency unchanged
/// even when the slide Part itself is restored through a prepared topology.
/// Reuse the native source loader and scene cache so semantic ceilings,
/// managed source data, MCE processing, and freshness fences stay identical to
/// an ordinary [`SourceSlide::layout`] read.
pub(super) fn layout_for_target(
    owner: &SourceInner,
    target: &PackURI,
    relationship_id: &str,
) -> Result<Layout> {
    let pending = with_fences(owner, || {
        let part = load_part(
            owner,
            target,
            CT_SLIDE_LAYOUT,
            crate::open_options::selected(owner.semantic_limits.as_ref())
                .shape()
                .input_bytes(),
        )?;
        prepare_layout(owner, part, relationship_id)
    })?;
    finish_layout(owner, pending)
}

impl SourceSlide {
    /// Whether the slide is authored as hidden.
    ///
    /// # Errors
    /// Returns typed source, execution, limit, relationship or XML errors.
    pub fn is_hidden(&self) -> Result<bool> {
        with_fences(&self.owner, || {
            let part = load_part(
                &self.owner,
                &self.data.part_uri,
                CT_SLIDE,
                crate::open_options::selected(self.owner.semantic_limits.as_ref())
                    .shape()
                    .input_bytes(),
            )?;
            SlidePart::from_part(&part)?.is_hidden()
        })
    }

    /// Acquire the slide's optional layout snapshot.
    ///
    /// # Errors
    /// Returns typed source, execution, limit, relationship or XML errors.
    pub fn layout(&self) -> Result<Option<Layout>> {
        let pending = with_fences(&self.owner, || {
            let Some((relationship_id, part)) = selected_relationship(
                &self.owner,
                self,
                SLIDE_LAYOUT_REL,
                "slideLayout",
                CT_SLIDE_LAYOUT,
                crate::open_options::selected(self.owner.semantic_limits.as_ref())
                    .shape()
                    .input_bytes(),
                "source slide has multiple layout relationships",
            )?
            else {
                return Ok(None);
            };

            prepare_layout(&self.owner, part, &relationship_id).map(Some)
        })?;
        pending
            .map(|pending| finish_layout(&self.owner, pending))
            .transpose()
    }

    /// Acquire the slide's optional notes snapshot.
    ///
    /// # Errors
    /// Returns typed source, execution, limit, relationship or XML errors.
    pub fn notes(&self) -> Result<Option<Notes>> {
        let pending = with_fences(&self.owner, || {
            let semantic = crate::open_options::selected(self.owner.semantic_limits.as_ref());
            let shape_limits = semantic.shape();
            let notes_limits = semantic.notes().constrained_by_source(self.owner.limits);
            let Some((relationship_id, part)) = selected_relationship(
                &self.owner,
                self,
                NOTES_SLIDE_REL,
                "notesSlide",
                CT_NOTES_SLIDE,
                notes_limits.max_input_bytes,
                "source slide has multiple notes relationships",
            )?
            else {
                return Ok(None);
            };

            notes::validate_with_limits(&self.owner, &self.data.part_uri, &part, notes_limits)?;
            let candidate = acquire_scene(&self.owner, &part, |source| {
                let parsed = crate::shape::Scene::read(source.blob(), shape_limits)?;
                let scene = OwnedScene::from_source(parsed, source.source_data().clone())?;
                strict_scene(&scene, "notes")?;
                Ok(scene)
            })?;
            let text = notes_text(&candidate.scene)?;
            Ok(Some((
                RelatedPart::new(part, &relationship_id),
                text,
                candidate,
            )))
        })?;
        let Some((related, text, candidate)) = pending else {
            return Ok(None);
        };
        let scene = publish_scene(&self.owner, candidate)?;
        Ok(Some(Notes::from_parts(related, text, scene)))
    }

    /// Acquire the slide's optional comments snapshot.
    ///
    /// # Errors
    /// Returns typed source, execution, limit, relationship or XML errors.
    pub fn comments(&self) -> Result<Option<Comments>> {
        with_fences(&self.owner, || {
            let comment_limits = crate::open_options::selected(self.owner.semantic_limits.as_ref())
                .comments()
                .constrained_by_source(self.owner.limits);
            let Some((relationship_id, part)) = selected_relationship(
                &self.owner,
                self,
                COMMENTS_REL,
                "comments",
                CT_COMMENTS,
                comment_limits.max_input_bytes,
                "source slide has multiple comments relationships",
            )?
            else {
                return Ok(None);
            };

            if !part.source_relationships().is_empty() {
                return Err(Error::Invalid(
                    "source comments part must not have relationships".into(),
                ));
            }
            let comments = parse_slide_comments_strict(part.blob(), comment_limits)?;
            if comments.len() > crate::comments::MAX_COMMENTS_PER_SLIDE {
                return Err(Error::Limit {
                    kind: crate::opened::LimitKind::Codec(
                        crate::opened::CodecLimitKind::CommentEntries,
                    ),
                    observed: comments.len(),
                    maximum: crate::comments::MAX_COMMENTS_PER_SLIDE,
                });
            }
            let authors = selected_source_authors(&self.owner, &comments, comment_limits)?;
            let (author_related, author_values) = match authors {
                Some((relationship_id, part, authors)) => (
                    Some(RelatedPart::new(part, &relationship_id)),
                    Some(authors),
                ),
                None => (None, None),
            };
            Comments::from_parts(
                RelatedPart::new(part, &relationship_id),
                author_related,
                comments,
                author_values.as_deref(),
            )
            .map(Some)
        })
    }

    /// Acquire all chart snapshots in lexical relationship-identifier order.
    ///
    /// # Errors
    /// Returns typed source, execution, limit, relationship or XML errors.
    pub fn charts(&self) -> Result<Vec<Chart>> {
        with_fences(&self.owner, || {
            let slide_part = self.owner.package.part(&self.data.part_uri)?;
            let relationships = chart_relationships(slide_part.rels())?;
            let mut charts = Vec::new();
            charts
                .try_reserve_exact(relationships.len())
                .map_err(|source| Error::Allocation {
                    resource: "source chart snapshots",
                    source,
                })?;

            for relationship in relationships {
                charts.push(read_chart(&self.owner, relationship)?);
            }
            Ok(charts)
        })
    }

    /// The chart-space XML of the ordinary or extended chart this slide relates through
    /// `relationship_id` (a chart frame's payload relationship), or `None`
    /// when the slide has no such chart relationship.
    ///
    /// # Errors
    /// Returns typed source, execution, limit, relationship or XML errors.
    pub fn chart_xml(&self, relationship_id: &str) -> Result<Option<Vec<u8>>> {
        with_fences(&self.owner, || {
            let slide_part = self.owner.package.part(&self.data.part_uri)?;
            let Some(relationship) = slide_part.rels().get(relationship_id) else {
                return Ok(None);
            };
            let extended =
                relationship.reltype() == litchi_drawingml::chart::chartex::RELATIONSHIP_TYPE;
            if !extended && !is_relationship_type(relationship.reltype(), CHART_REL, "chart") {
                return Ok(None);
            }
            if relationship.target_mode() == TargetMode::External {
                return Err(Error::Invalid(
                    "source chart relationship must be internal".into(),
                ));
            }
            let target = relationship.target_partname()?;
            let content_type = if extended {
                litchi_drawingml::chart::chartex::CONTENT_TYPE
            } else {
                CT_CHART
            };
            let part = load_part(&self.owner, &target, content_type, MAX_CHART_XML)?;
            Ok(Some(part.blob().to_vec()))
        })
    }

    /// The Office Math equations of this slide's text (record 0321); see
    /// [`crate::slide::Slide::equations`].
    ///
    /// # Errors
    /// Returns typed source, execution, limit or XML errors.
    pub fn equations(&self) -> Result<Vec<crate::shape::math::Equation>> {
        with_fences(&self.owner, || {
            // The slide's own bytes: the equations sit in a
            // markup-compatibility choice the processed part drops.
            let view = self.owner.package.part(&self.data.part_uri)?;
            let declared = view.declared_uncompressed_size()?;
            let maximum = crate::open_options::selected(self.owner.semantic_limits.as_ref())
                .shape()
                .input_bytes();
            if declared > maximum as u64 {
                return Err(Error::Limit {
                    kind: crate::opened::LimitKind::Codec(
                        crate::opened::CodecLimitKind::SourceBytes,
                    ),
                    observed: usize::try_from(declared).unwrap_or(usize::MAX),
                    maximum,
                });
            }
            let data = view.data()?;
            crate::source_limits::validate_xml_limits(data.as_bytes(), self.owner.limits)?;
            crate::shape::math::equations(data.as_bytes())
        })
    }

    /// Resolve a click action of this slide's shapes or text through the
    /// slide's relationships (record 0321); see
    /// [`crate::slide::Slide::hyperlink`].
    ///
    /// # Errors
    /// Returns typed source, execution or relationship errors.
    pub fn hyperlink(
        &self,
        click: &crate::hyperlinks::Click,
    ) -> Result<Option<crate::hyperlinks::Link>> {
        with_fences(&self.owner, || {
            let slide_part = self.owner.package.part(&self.data.part_uri)?;
            let rels = slide_part.rels();
            crate::hyperlinks::Link::resolve(
                click,
                |id| {
                    let Some(relationship) = rels.get(id) else {
                        return Ok(None);
                    };
                    Ok(Some(
                        if relationship.target_mode() == TargetMode::External {
                            (true, relationship.target_ref().to_owned())
                        } else {
                            (false, relationship.target_partname()?.to_string())
                        },
                    ))
                },
                |part| {
                    Ok(self
                        .owner
                        .slides
                        .iter()
                        .position(|slide| slide.part_uri.as_str() == part)
                        .map(Position::new))
                },
            )
        })
    }

    /// The data-model XML of the SmartArt this slide relates through
    /// `relationship_id` (a SmartArt frame's `dgm:relIds/@r:dm`), or `None`
    /// when the slide has no such diagram-data relationship (record 0284).
    ///
    /// # Errors
    /// Returns typed source, execution, limit, relationship or XML errors.
    pub fn diagram_xml(&self, relationship_id: &str) -> Result<Option<Vec<u8>>> {
        with_fences(&self.owner, || {
            let slide_part = self.owner.package.part(&self.data.part_uri)?;
            let Some(relationship) = slide_part.rels().get(relationship_id) else {
                return Ok(None);
            };
            if !is_relationship_type(
                relationship.reltype(),
                litchi_drawingml::diagram::DIAGRAM_DATA_REL,
                "diagramData",
            ) {
                return Ok(None);
            }
            if relationship.target_mode() == TargetMode::External {
                return Err(Error::Invalid(
                    "source diagram relationship must be internal".into(),
                ));
            }
            let target = relationship.target_partname()?;
            let part = load_part(
                &self.owner,
                &target,
                litchi_opc::constants::content_type::DML_DIAGRAM_DATA,
                litchi_drawingml::diagram::edit::MAX_PART_BYTES,
            )?;
            Ok(Some(part.blob().to_vec()))
        })
    }

    /// Acquire the chart at a positional relationship index.
    ///
    /// # Errors
    /// Returns typed source, execution, limit, relationship or XML errors.
    pub fn chart(&self, position: Position) -> Result<Option<Chart>> {
        with_fences(&self.owner, || {
            let slide_part = self.owner.package.part(&self.data.part_uri)?;
            let relationships = chart_relationships(slide_part.rels())?;
            relationships
                .get(position.get())
                .map(|relationship| read_chart(&self.owner, relationship).map(Some))
                .unwrap_or(Ok(None))
        })
    }
}
