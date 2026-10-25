//! Borrowed semantic slide, layout, and master views.
//!
//! These types retain only a package borrow and a validated low-level part
//! view. XML decoding, shape indexing, and relationship traversal are
//! delegated to the sibling [`super::codec`] and [`super::package`] layers so
//! the public API stays contextual without owning duplicate buffers.

use litchi_opc::OpcPackage;

use crate::Result;
use crate::parts::{SlideLayoutPart, SlideMasterPart, SlidePart};
use crate::shape::Scene;

use super::{codec, package};

/// Semantic slide selection by exact producer-visible name or checked position.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Selector<'a> {
    /// Select by exact producer-visible name; ambiguous names are refused.
    Name(std::borrow::Cow<'a, str>),
    /// Select by zero-based presentation order.
    Position(litchi_core::Position),
}

impl<'a> From<&'a str> for Selector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(value.into())
    }
}
impl From<String> for Selector<'_> {
    fn from(value: String) -> Self {
        Self::Name(value.into())
    }
}
impl<'a> From<&'a String> for Selector<'a> {
    fn from(value: &'a String) -> Self {
        Self::Name(value.as_str().into())
    }
}
impl From<litchi_core::Position> for Selector<'_> {
    fn from(value: litchi_core::Position) -> Self {
        Self::Position(value)
    }
}
impl From<usize> for Selector<'_> {
    fn from(value: usize) -> Self {
        Self::Position(litchi_core::Position::new(value))
    }
}

/// A slide master by its typed `p:sldMasterId/@id`, unique in the
/// presentation (record 0356a).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct MasterId(pub u32);

/// A slide layout by its typed `p:sldLayoutId/@id`, unique in the
/// presentation across every master (record 0356a).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct LayoutId(pub u32);

/// The owner of slide-scoped markup: a slide, a slide layout or a slide
/// master (record 0356a). Families that edit markup found in all three
/// (color maps, placeholders) take it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Owner<'a> {
    /// A slide, by name or position.
    Slide(Selector<'a>),
    /// A slide layout, by its layout ID.
    Layout(LayoutId),
    /// A slide master, by its master ID.
    Master(MasterId),
}

impl<'a> From<Selector<'a>> for Owner<'a> {
    fn from(value: Selector<'a>) -> Self {
        Self::Slide(value)
    }
}
impl<'a> From<&'a str> for Owner<'a> {
    fn from(value: &'a str) -> Self {
        Self::Slide(value.into())
    }
}
impl From<usize> for Owner<'_> {
    fn from(value: usize) -> Self {
        Self::Slide(value.into())
    }
}
impl From<litchi_core::Position> for Owner<'_> {
    fn from(value: litchi_core::Position) -> Self {
        Self::Slide(value.into())
    }
}
impl From<LayoutId> for Owner<'_> {
    fn from(value: LayoutId) -> Self {
        Self::Layout(value)
    }
}
impl From<MasterId> for Owner<'_> {
    fn from(value: MasterId) -> Self {
        Self::Master(value)
    }
}

/// Contextual semantic view of one slide.
pub struct Slide<'a> {
    package: &'a OpcPackage,
    part: SlidePart<'a>,
    pub(crate) semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
    pub(crate) shape_limits: crate::shape::Limits,
}

impl<'a> Slide<'a> {
    /// Construct a slide view from its package and validated part view.
    #[must_use]
    pub fn new(package: &'a OpcPackage, part: SlidePart<'a>) -> Self {
        Self::new_with_semantic_limits(package, part, None)
    }

    pub(crate) fn new_with_semantic_limits(
        package: &'a OpcPackage,
        part: SlidePart<'a>,
        semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
    ) -> Self {
        Self {
            package,
            part,
            shape_limits: crate::open_options::selected(semantic_limits.as_ref()).shape(),
            semantic_limits,
        }
    }

    pub(crate) fn new_with_shape_and_semantic_limits(
        package: &'a OpcPackage,
        part: SlidePart<'a>,
        shape_limits: crate::shape::Limits,
        semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
    ) -> Self {
        Self {
            package,
            part,
            shape_limits,
            semantic_limits,
        }
    }

    pub(crate) fn semantic_limits(&self) -> crate::open_options::SemanticLimits {
        crate::open_options::selected(self.semantic_limits.as_ref())
    }

    /// Borrow the underlying validated slide part view.
    #[inline]
    #[must_use]
    pub fn part(&self) -> &SlidePart<'a> {
        &self.part
    }

    /// Return the producer name, falling back to the OPC part name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn name(&self) -> Result<String> {
        codec::slide_name(&self.part)
    }

    /// Whether the slide is hidden in the presentation graph.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn is_hidden(&self) -> Result<bool> {
        codec::slide_hidden(&self.part)
    }

    /// Flatten `DrawingML` text runs in source order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn text(&self) -> Result<String> {
        codec::slide_text(&self.part)
    }

    /// [`Self::text`], appending the repairs the extraction applied to
    /// `tolerance`, stamped with the slide's part name (see
    /// [`crate::parts::SlidePart::text_with_tolerance`]).
    ///
    /// # Errors
    ///
    /// Returns the errors [`Self::text`] returns.
    pub fn text_with_tolerance(
        &self,
        tolerance: &mut crate::leniency::ToleranceReport,
    ) -> Result<String> {
        self.part.text_with_tolerance(tolerance)
    }

    /// Build the bounded borrowed scene for the slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shapes(&self) -> Result<Scene<'a>> {
        codec::slide_shapes_with_limits(&self.part, self.shape_limits)
    }

    /// Retain an independently owned shape scene sharing this part's payload.
    /// Handles from the scene remain valid after the package and slide view drop.
    ///
    /// # Errors
    /// Returns the same validation and parse errors as [`Self::shapes`].
    pub fn scene(&self) -> Result<crate::shape::owned::Scene> {
        crate::shape::owned::Scene::from_parsed(
            self.part.shapes_with_limits(self.shape_limits)?,
            self.part.part().blob_arc(),
        )
    }

    /// Parse this slide's animation timing tree and build list.
    ///
    /// Stale cross-references are repaired the way `PowerPoint` repairs them:
    /// an animation, interactive trigger, or build entry whose target shape
    /// no longer exists on the slide (or cannot host the entry) is skipped
    /// while the rest of the tree plays, and a shape ID carried by two shapes
    /// resolves to the first one. Each repair is recorded, with this slide's
    /// part name and the shape ID, in the returned sequence's
    /// [`crate::animations::Sequence::tolerance_report`]; an empty report
    /// means the timing tree conformed. A slide without `p:timing` yields an
    /// empty sequence.
    ///
    /// # Errors
    ///
    /// Returns an error if the slide XML is malformed, exceeds a read limit,
    /// or carries a structurally invalid timing tree, or if a surviving build
    /// entry's host relationship cannot be resolved inside the package.
    pub fn animations(&self) -> Result<crate::animations::Sequence> {
        crate::animations::parse_package_slide_with_limits(
            self.package,
            self.part.part().partname(),
            self.semantic_limits()
                .animations()
                .constrained_by_source(self.package.read_limits()),
        )
    }

    /// Read the optional direct programmable-tag list attached to this slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn tags(&self) -> Result<Option<crate::tag::List>> {
        package::slide_tags(
            self.package,
            &self.part,
            self.semantic_limits()
                .tag()
                .constrained_by_source(self.package.read_limits()),
        )
    }

    /// Read the optional programmable-tag list attached to one semantic shape.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_tags<'k>(
        &self,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::tag::List>> {
        package::slide_shape_tags(
            self.package,
            &self.part,
            shape,
            self.semantic_limits()
                .tag()
                .constrained_by_source(self.package.read_limits()),
            self.shape_limits,
        )
    }

    /// Read the optional typed classification outcome attached to one shape.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_classification<'k>(
        &self,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::shape::classification::Snapshot>> {
        package::slide_shape_classification(self.package, &self.part, shape, self.shape_limits)
    }

    /// Capture this slide's inert `[MS-PPTX]` change-tracking identifiers.
    ///
    /// # Errors
    ///
    /// Returns an error when the bounded extension XML is malformed. Duplicate
    /// producer values remain readable for checked-position repair.
    pub fn change_tracking(&self) -> Result<crate::change_tracking::Snapshot> {
        crate::change_tracking::load_with_limits(
            self.package,
            self.part.part().partname(),
            self.shape_limits,
        )
    }

    /// Read the optional typed `p15:designElem` value attached to one shape.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_design_element<'k>(
        &self,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::shape::designer::Snapshot>> {
        package::slide_shape_design_element(self.package, &self.part, shape, self.shape_limits)
    }

    /// Read one shape's source-bound Designer drawing properties.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_designer_properties<'k>(
        &self,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<crate::shape::designer::properties::Snapshot> {
        self.shape_designer_properties_with_limits(shape, crate::shape::designer::Limits::default())
    }

    /// Read shape Designer properties under caller-supplied resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn shape_designer_properties_with_limits<'k>(
        &self,
        shape: impl Into<crate::shape::Key<'k>>,
        limits: crate::shape::designer::Limits,
    ) -> Result<crate::shape::designer::properties::Snapshot> {
        package::slide_shape_designer_properties(
            self.package,
            &self.part,
            shape,
            limits,
            self.shape_limits,
        )
    }

    /// Read the source-bound Designer tags owned by this slide's stable ID.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn designer_tags(
        &self,
    ) -> Result<crate::presentation_properties::metadata::designer_tags::Snapshot> {
        self.designer_tags_with_limits(crate::shape::designer::Limits::default())
    }

    /// Read this slide's Designer tags under caller-supplied resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn designer_tags_with_limits(
        &self,
        limits: crate::shape::designer::Limits,
    ) -> Result<crate::presentation_properties::metadata::designer_tags::Snapshot> {
        package::slide_designer_tags(self.package, &self.part, limits)
    }

    /// Read all contextual 3D-model owners on this slide in source order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn model3ds(&self) -> Result<Vec<crate::model3d::Model>> {
        let limits = crate::model3d::codec::Limits::from_source_limits(self.package.read_limits());
        crate::model3d::package::load_all(self.package, self.part.part().partname(), limits)
    }

    /// Read the model3d owner attached to one semantic shape, if present.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn model3d<'k>(
        &self,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::model3d::Model>> {
        let limits = crate::model3d::codec::Limits::from_source_limits(self.package.read_limits());
        crate::model3d::package::load(
            self.package,
            self.part.part().partname(),
            shape.into(),
            limits,
        )
    }

    /// Inspect all tag relationships on this slide in stable relationship-ID
    /// order. Shape-owned and unanchored producer markup remains visible here
    /// but is not flattened into [`Self::tags`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn tag_inventory(&self) -> Result<Vec<crate::tag::Source>> {
        package::slide_tag_inventory(
            self.package,
            &self.part,
            self.semantic_limits()
                .tag()
                .constrained_by_source(self.package.read_limits()),
        )
    }

    /// Resolve the ordinary charts attached to this slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn charts(&self) -> Result<Vec<crate::chart::Part<'a>>> {
        package::slide_charts(self.package, &self.part)
    }

    /// The chart-space XML of the ordinary (`c:chartSpace`) or extended
    /// (`cx:chartSpace`) chart this slide relates through
    /// `relationship_id` (a chart frame's payload relationship), or `None`
    /// when the slide has no such chart relationship.
    ///
    /// # Errors
    ///
    /// Returns an error for an external relationship or a target that is
    /// not a chart part.
    pub fn chart_xml(&self, relationship_id: &str) -> Result<Option<&'a [u8]>> {
        let Some(relationship) = self.part.part().rels().get(relationship_id) else {
            return Ok(None);
        };
        let extended =
            relationship.reltype() == litchi_drawingml::chart::chartex::RELATIONSHIP_TYPE;
        if !extended
            && !crate::parts::is_relationship_type(
                relationship.reltype(),
                litchi_opc::constants::relationship_type::CHART,
                "chart",
            )
        {
            return Ok(None);
        }
        if relationship.is_external() {
            return Err(crate::Error::Relationship(
                "chart relationships cannot be external".to_owned(),
            ));
        }
        let target = relationship.target_partname()?;
        if extended {
            let part = self.package.get_part(&target)?;
            if part.content_type() != litchi_drawingml::chart::chartex::CONTENT_TYPE {
                return Err(crate::Error::Relationship(
                    "an extended chart relationship targets a part that is not a chartex part"
                        .to_owned(),
                ));
            }
            return Ok(Some(part.blob()));
        }
        let part = crate::chart::Part::from_part(self.package.get_part(&target)?)?;
        Ok(Some(part.part().blob()))
    }

    /// The Office Math equations of this slide's text, in document order
    /// (record 0321), read from the slide's own XML (the shape scene reads
    /// their picture fallbacks).
    ///
    /// # Errors
    ///
    /// Returns an error for malformed slide XML or exceeded bounds.
    pub fn equations(&self) -> Result<Vec<crate::shape::math::Equation>> {
        crate::shape::math::equations(self.part.part().blob())
    }

    /// Resolve a click action of this slide's shapes or text through the
    /// slide's relationships (record 0321): an external URL, a slide of the
    /// presentation, or a relationship-free action. `None` when the click
    /// names a relationship the slide lacks, or a part that is not a slide.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed relationship target or
    /// presentation slide list.
    pub fn hyperlink(
        &self,
        click: &crate::hyperlinks::Click,
    ) -> Result<Option<crate::hyperlinks::Link>> {
        let rels = self.part.part().rels();
        let package = self.package;
        crate::hyperlinks::Link::resolve(
            click,
            |id| {
                let Some(relationship) = rels.get(id) else {
                    return Ok(None);
                };
                Ok(Some(if relationship.is_external() {
                    (true, relationship.target_ref().to_owned())
                } else {
                    (false, relationship.target_partname()?.to_string())
                }))
            },
            |part| {
                let presentation = crate::parts::PresentationPart::from_package(package)?;
                let rels = presentation.part().rels();
                for (index, reference) in presentation.slide_references()?.iter().enumerate() {
                    if let Some(relationship) = rels.get(reference.relationship_id())
                        && !relationship.is_external()
                        && relationship.target_partname()?.as_str() == part
                    {
                        return Ok(Some(litchi_core::Position::new(index)));
                    }
                }
                Ok(None)
            },
        )
    }

    /// The data-model XML of the SmartArt this slide relates through
    /// `relationship_id` (a SmartArt frame's `dgm:relIds/@r:dm`), or `None`
    /// when the slide has no such diagram-data relationship (record 0284).
    ///
    /// # Errors
    ///
    /// Returns an error for an external relationship or a target that is
    /// not a diagram data part.
    pub fn diagram_xml(&self, relationship_id: &str) -> Result<Option<&'a [u8]>> {
        let Some(relationship) = self.part.part().rels().get(relationship_id) else {
            return Ok(None);
        };
        if !crate::parts::is_relationship_type(
            relationship.reltype(),
            litchi_drawingml::diagram::DIAGRAM_DATA_REL,
            "diagramData",
        ) {
            return Ok(None);
        }
        if relationship.is_external() {
            return Err(crate::Error::Relationship(
                "diagram relationships cannot be external".to_owned(),
            ));
        }
        let target = relationship.target_partname()?;
        let part = self.package.get_part(&target)?;
        if part.content_type() != litchi_opc::constants::content_type::DML_DIAGRAM_DATA {
            return Err(crate::Error::Relationship(
                "the diagram relationship does not target a diagram data part".to_owned(),
            ));
        }
        Ok(Some(part.blob()))
    }

    /// Resolve Microsoft `ChartEx` parts attached to this slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn chart_extensions(&self) -> Result<Vec<crate::chart::extension::Part<'a>>> {
        package::slide_chart_extensions(self.package, &self.part)
    }

    /// Resolve the optional legacy comments list attached to this slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn comments(&self) -> Result<Option<crate::comments::ListPart<'a>>> {
        package::slide_comments(self.package, &self.part)
    }

    /// Resolve the optional legacy comments part into a detached semantic
    /// snapshot.  Only this slide's comments relationship and, when present,
    /// the presentation-level comment-author part are read; the package-wide
    /// comments graph is not materialized.
    ///
    /// # Errors
    ///
    /// Returns an error when the selected relationships, content types,
    /// author closure, or comments XML are malformed.
    pub fn comment_snapshot(&self) -> Result<Option<crate::presentation::related::Comments>> {
        package::slide_comment_snapshot(
            self.package,
            &self.part,
            self.semantic_limits()
                .comments()
                .constrained_by_source(self.package.read_limits()),
        )
    }

    /// Resolve the optional modern-comment list attached to this slide.
    ///
    /// Only this slide's modern-comment part and, when required, the
    /// presentation-level author part are read. The package-wide modern
    /// comment graph is not materialized.
    ///
    /// # Errors
    ///
    /// Returns an error if the selected relationships, content types, author
    /// closure, or modern-comment XML are malformed.
    pub fn modern_comments(&self) -> Result<Option<crate::presentation::modern_comments::List>> {
        crate::presentation::modern_comments::acquire_eager(self.package, &self.part)
    }

    /// Read this slide's optional speaker-notes part as a detached snapshot.
    ///
    /// The selected notes relationship and its bounded owning-slide,
    /// notes-master, theme, and presentation-root closure are read. The
    /// presentation-wide notes graph and other slides are not traversed. The
    /// returned snapshot retains its owned scene and remains usable after the
    /// package and slide view are dropped.
    ///
    /// # Errors
    ///
    /// Returns an error when the relationship is duplicated, external, or
    /// points to a missing, mistyped, malformed, or non-strict notes part.
    pub fn notes(&self) -> Result<Option<crate::presentation::related::Notes>> {
        package::slide_notes(
            self.package,
            &self.part,
            self.shape_limits,
            self.semantic_limits()
                .notes()
                .constrained_by_source(self.package.read_limits()),
        )
    }

    /// Read typed section, slide, and summary zoom metadata in this slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn zooms(&self) -> Result<crate::shape::zoom::Owner> {
        package::slide_zooms(self.package, &self.part)
    }

    /// Count semantic shapes in the slide's scene.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_count(&self) -> Result<usize> {
        Ok(self.shapes()?.len())
    }

    /// Return this slide's optional slide-library synchronization metadata.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_sync(
        &self,
    ) -> Result<Option<crate::presentation_properties::metadata::slide_sync::Properties>> {
        package::slide_sync(
            self.package,
            &self.part,
            crate::open_options::selected(self.semantic_limits.as_ref())
                .slide_sync()
                .constrained_by_source(self.package.read_limits()),
        )
    }

    /// Resolve the slide's optional layout in package context.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn layout(&self) -> Result<Option<SlideLayout<'a>>> {
        package::slide_layout(self.package, &self.part).map(|part| {
            part.map(|part| {
                SlideLayout::new_with_shape_limits(self.package, part, self.shape_limits)
            })
        })
    }
}

/// Contextual semantic view of one slide layout.
pub struct SlideLayout<'a> {
    package: &'a OpcPackage,
    part: SlideLayoutPart<'a>,
    shape_limits: crate::shape::Limits,
}

impl<'a> SlideLayout<'a> {
    /// Construct a layout view from its package and validated part view.
    #[must_use]
    pub fn new(package: &'a OpcPackage, part: SlideLayoutPart<'a>) -> Self {
        Self::new_with_shape_limits(package, part, crate::shape::Limits::DEFAULT)
    }

    pub(crate) fn new_with_shape_limits(
        package: &'a OpcPackage,
        part: SlideLayoutPart<'a>,
        shape_limits: crate::shape::Limits,
    ) -> Self {
        Self {
            package,
            part,
            shape_limits,
        }
    }

    /// Borrow the underlying validated layout part view.
    #[inline]
    #[must_use]
    pub fn part(&self) -> &SlideLayoutPart<'a> {
        &self.part
    }

    /// The layout's typed `p:sldLayoutId/@id`, the selector the edit
    /// families take (record 0356a).
    ///
    /// # Errors
    ///
    /// Returns an error when no master lists this layout with an ID, or for
    /// a malformed master graph.
    pub fn id(&self) -> Result<LayoutId> {
        crate::master_layout::package::layout_id_of(self.package, self.part.part().partname())?
            .map(LayoutId)
            .ok_or(crate::Error::refused(
                litchi_core::edit::RefusalKind::NotFound,
                "layout_id",
                "no slide master lists this layout with an ID",
            ))
    }

    /// Return the producer name, falling back to the OPC part name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn name(&self) -> Result<String> {
        codec::layout_name(&self.part)
    }

    /// Return the optional `PresentationML` layout kind token.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn kind(&self) -> Result<Option<String>> {
        codec::layout_kind(&self.part)
    }

    /// Build the bounded borrowed shape scene for this layout.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shapes(&self) -> Result<Scene<'a>> {
        codec::layout_shapes_with_limits(&self.part, self.shape_limits)
    }

    /// Retain an independently owned shape scene sharing this part's payload.
    /// Handles from the scene remain valid after the package and slide view drop.
    ///
    /// # Errors
    /// Returns the same validation and parse errors as [`Self::shapes`].
    pub fn scene(&self) -> Result<crate::shape::owned::Scene> {
        crate::shape::owned::Scene::from_parsed(
            self.part.shapes_with_limits(self.shape_limits)?,
            self.part.part().blob_arc(),
        )
    }

    /// Read the optional theme override attached to this layout.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn theme_override(&self) -> Result<Option<crate::shape::theme::Override>> {
        package::layout_theme_override(self.package, &self.part)
    }

    /// Resolve the required master in package context.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn master(&self) -> Result<SlideMaster<'a>> {
        package::layout_master(self.package, &self.part)
            .map(|part| SlideMaster::new_with_shape_limits(self.package, part, self.shape_limits))
    }
}

/// Contextual semantic view of one slide master.
pub struct SlideMaster<'a> {
    package: &'a OpcPackage,
    part: SlideMasterPart<'a>,
    shape_limits: crate::shape::Limits,
}

impl<'a> SlideMaster<'a> {
    /// Construct a master view from its package and validated part view.
    #[must_use]
    pub fn new(package: &'a OpcPackage, part: SlideMasterPart<'a>) -> Self {
        Self::new_with_shape_limits(package, part, crate::shape::Limits::DEFAULT)
    }

    pub(crate) fn new_with_shape_limits(
        package: &'a OpcPackage,
        part: SlideMasterPart<'a>,
        shape_limits: crate::shape::Limits,
    ) -> Self {
        Self {
            package,
            part,
            shape_limits,
        }
    }

    /// Borrow the underlying validated master part view.
    #[inline]
    #[must_use]
    pub fn part(&self) -> &SlideMasterPart<'a> {
        &self.part
    }

    /// The master's typed `p:sldMasterId/@id`, the selector the edit
    /// families take (record 0356a).
    ///
    /// # Errors
    ///
    /// Returns an error when the presentation does not list this master, or
    /// for a malformed presentation part.
    pub fn id(&self) -> Result<MasterId> {
        crate::master_layout::package::master_id_of(self.package, self.part.part().partname())?
            .map(MasterId)
            .ok_or(crate::Error::refused(
                litchi_core::edit::RefusalKind::NotFound,
                "master_id",
                "the presentation does not list this master",
            ))
    }

    /// Return the producer name, falling back to the OPC part name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn name(&self) -> Result<String> {
        codec::master_name(&self.part)
    }

    /// Whether `PowerPoint` should preserve the master after editing.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn is_preserved(&self) -> Result<bool> {
        codec::master_preserved(&self.part)
    }

    /// Build the bounded borrowed shape scene for this master.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shapes(&self) -> Result<Scene<'a>> {
        codec::master_shapes_with_limits(&self.part, self.shape_limits)
    }

    /// Retain an independently owned shape scene sharing this part's payload.
    /// Handles from the scene remain valid after the package and slide view drop.
    ///
    /// # Errors
    /// Returns the same validation and parse errors as [`Self::shapes`].
    pub fn scene(&self) -> Result<crate::shape::owned::Scene> {
        crate::shape::owned::Scene::from_parsed(
            self.part.shapes_with_limits(self.shape_limits)?,
            self.part.part().blob_arc(),
        )
    }

    /// Read the theme reached from this slide master.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn theme(&self) -> Result<Option<crate::shape::theme::ThemeSummary>> {
        package::master_theme(self.package, &self.part)
    }

    /// Resolve all layouts reachable from this master in XML order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn layouts(&self) -> Result<Vec<SlideLayout<'a>>> {
        package::master_layouts(self.package, &self.part).map(|parts| {
            parts
                .into_iter()
                .map(|part| {
                    SlideLayout::new_with_shape_limits(self.package, part, self.shape_limits)
                })
                .collect()
        })
    }
}
