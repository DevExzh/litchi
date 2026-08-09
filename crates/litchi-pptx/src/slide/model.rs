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

/// Checked selector for an ordered slide graph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Key<'a> {
    /// Select by exact producer-visible name.
    Name(&'a str),
    /// Select by zero-based presentation order.
    Index(usize),
}

impl<'a> From<&'a str> for Key<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(value)
    }
}

impl From<usize> for Key<'_> {
    fn from(value: usize) -> Self {
        Self::Index(value)
    }
}

/// Contextual semantic view of one slide.
pub struct Slide<'a> {
    package: &'a OpcPackage,
    part: SlidePart<'a>,
}

impl<'a> Slide<'a> {
    /// Construct a slide view from its package and validated part view.
    #[must_use]
    pub fn new(package: &'a OpcPackage, part: SlidePart<'a>) -> Self {
        Self { package, part }
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

    /// Build the bounded borrowed scene for the slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shapes(&self) -> Result<Scene<'a>> {
        codec::slide_shapes(&self.part)
    }

    /// Read the optional direct programmable-tag list attached to this slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn tags(&self) -> Result<Option<crate::tag::List>> {
        package::slide_tags(self.package, &self.part)
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
        package::slide_shape_tags(self.package, &self.part, shape)
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
        package::slide_shape_classification(self.package, &self.part, shape)
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
        package::slide_shape_design_element(self.package, &self.part, shape)
    }

    /// Read one shape's source-bound Designer drawing properties.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_designer_properties<'k>(
        &self,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<crate::shape::designer::PropertiesSnapshot> {
        self.shape_designer_properties_with_limits(shape, crate::shape::designer::Limits::default())
    }

    /// Read shape Designer properties under caller-supplied resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_designer_properties_with_limits<'k>(
        &self,
        shape: impl Into<crate::shape::Key<'k>>,
        limits: crate::shape::designer::Limits,
    ) -> Result<crate::shape::designer::PropertiesSnapshot> {
        package::slide_shape_designer_properties(self.package, &self.part, shape, limits)
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
    pub fn designer_tags_with_limits(
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
        crate::model3d::package::load_all(self.package, self.part.part().partname())
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
        crate::model3d::package::load(self.package, self.part.part().partname(), shape.into())
    }

    /// Inspect all tag relationships on this slide in stable relationship-ID
    /// order. Shape-owned and unanchored producer markup remains visible here
    /// but is not flattened into [`Self::tags`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn tag_inventory(&self) -> Result<Vec<crate::tag::Source>> {
        package::slide_tag_inventory(self.package, &self.part)
    }

    /// Resolve the ordinary charts attached to this slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn charts(&self) -> Result<Vec<crate::chart::Part<'a>>> {
        package::slide_charts(self.package, &self.part)
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
        package::slide_sync(self.package, &self.part)
    }

    /// Resolve the slide's optional layout in package context.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn layout(&self) -> Result<Option<SlideLayout<'a>>> {
        package::slide_layout(self.package, &self.part)
            .map(|part| part.map(|part| SlideLayout::new(self.package, part)))
    }
}

/// Contextual semantic view of one slide layout.
pub struct SlideLayout<'a> {
    package: &'a OpcPackage,
    part: SlideLayoutPart<'a>,
}

impl<'a> SlideLayout<'a> {
    /// Construct a layout view from its package and validated part view.
    #[must_use]
    pub fn new(package: &'a OpcPackage, part: SlideLayoutPart<'a>) -> Self {
        Self { package, part }
    }

    /// Borrow the underlying validated layout part view.
    #[inline]
    #[must_use]
    pub fn part(&self) -> &SlideLayoutPart<'a> {
        &self.part
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
        codec::layout_shapes(&self.part)
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
            .map(|part| SlideMaster::new(self.package, part))
    }
}

/// Contextual semantic view of one slide master.
pub struct SlideMaster<'a> {
    package: &'a OpcPackage,
    part: SlideMasterPart<'a>,
}

impl<'a> SlideMaster<'a> {
    /// Construct a master view from its package and validated part view.
    #[must_use]
    pub fn new(package: &'a OpcPackage, part: SlideMasterPart<'a>) -> Self {
        Self { package, part }
    }

    /// Borrow the underlying validated master part view.
    #[inline]
    #[must_use]
    pub fn part(&self) -> &SlideMasterPart<'a> {
        &self.part
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
        codec::master_shapes(&self.part)
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
                .map(|part| SlideLayout::new(self.package, part))
                .collect()
        })
    }
}
