//! Contextual slide, layout, and master facades.

use litchi_opc::OpcPackage;

use crate::Result;
use crate::parts::{SlideLayoutPart, SlideMasterPart, SlidePart};
use crate::shape::Scene;

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
    /// Construct a slide view from its package and low-level part view.
    pub fn new(package: &'a OpcPackage, part: SlidePart<'a>) -> Self {
        Self { package, part }
    }

    /// Borrow the underlying slide part view.
    #[inline]
    pub fn part(&self) -> &SlidePart<'a> {
        &self.part
    }

    /// Return the producer name, falling back to the OPC part name.
    pub fn name(&self) -> Result<String> {
        self.part.name()
    }

    /// Whether the slide is hidden in the presentation graph.
    pub fn is_hidden(&self) -> Result<bool> {
        self.part.is_hidden()
    }

    /// Flatten DrawingML text runs in source order.
    pub fn text(&self) -> Result<String> {
        self.part.text()
    }

    /// Build the bounded borrowed scene for the slide.
    pub fn shapes(&self) -> Result<Scene<'a>> {
        self.part.shapes()
    }

    /// Read the optional direct programmable-tag list attached to this slide.
    pub fn tags(&self) -> Result<Option<crate::tag::List>> {
        Ok(crate::tag::load(self.package, self.part.part().partname())?
            .map(crate::tag::Source::into_list))
    }

    /// Read the optional programmable-tag list attached to one semantic shape.
    pub fn shape_tags<'k>(
        &self,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::tag::List>> {
        Ok(
            crate::tag::shape::load(self.package, self.part.part().partname(), shape)?
                .map(crate::tag::Source::into_list),
        )
    }

    /// Inspect all tag relationships on this slide in stable relationship-ID
    /// order. Shape-owned and unanchored producer markup remains visible here
    /// but is not flattened into [`Self::tags`].
    pub fn tag_inventory(&self) -> Result<Vec<crate::tag::Source>> {
        crate::tag::discover(self.part.part(), self.package).map_err(Into::into)
    }

    /// Resolve the ordinary charts attached to this slide.
    pub fn charts(&self) -> Result<Vec<crate::chart::Part<'a>>> {
        self.part.charts(self.package)
    }

    /// Resolve Microsoft ChartEx parts attached to this slide.
    pub fn chart_extensions(&self) -> Result<Vec<crate::chart::extension::Part<'a>>> {
        self.part.chart_extensions(self.package)
    }

    /// Resolve the optional legacy comments list attached to this slide.
    pub fn comments(&self) -> Result<Option<crate::comments::ListPart<'a>>> {
        self.part.comments(self.package)
    }

    /// Count semantic shapes in the slide's scene.
    pub fn shape_count(&self) -> Result<usize> {
        Ok(self.shapes()?.len())
    }

    /// Return this slide's optional slide-library synchronization metadata.
    pub fn slide_sync(
        &self,
    ) -> Result<Option<crate::presentation_properties::metadata::slide_sync::Properties>> {
        let part_name = self.part.part().partname();
        let mut matches = crate::presentation_properties::metadata::slide_sync::load(self.package)?
            .into_iter()
            .filter(|entry| entry.slide_part_name == *part_name);
        Ok(matches.next().map(|entry| entry.properties))
    }

    /// Resolve the slide's optional layout in package context.
    pub fn layout(&self) -> Result<Option<SlideLayout<'a>>> {
        let part = self.part.layout(self.package)?;
        Ok(part.map(|part| SlideLayout::new(self.package, part)))
    }
}

/// Contextual semantic view of one slide layout.
pub struct SlideLayout<'a> {
    package: &'a OpcPackage,
    part: SlideLayoutPart<'a>,
}

impl<'a> SlideLayout<'a> {
    /// Construct a layout view from its package and low-level part view.
    pub fn new(package: &'a OpcPackage, part: SlideLayoutPart<'a>) -> Self {
        Self { package, part }
    }

    /// Borrow the underlying layout part view.
    #[inline]
    pub fn part(&self) -> &SlideLayoutPart<'a> {
        &self.part
    }

    /// Return the producer name, falling back to the OPC part name.
    pub fn name(&self) -> Result<String> {
        self.part.name()
    }

    /// Return the optional PresentationML layout kind token.
    pub fn kind(&self) -> Result<Option<String>> {
        self.part.kind()
    }

    /// Build the bounded borrowed scene for the layout.
    pub fn shapes(&self) -> Result<Scene<'a>> {
        self.part.shapes()
    }

    /// Read the optional theme override attached to this layout.
    pub fn theme_override(&self) -> Result<Option<crate::shape::theme::Override>> {
        self.part.theme_override(self.package)
    }

    /// Resolve the required master in package context.
    pub fn master(&self) -> Result<SlideMaster<'a>> {
        Ok(SlideMaster::new(
            self.package,
            self.part.master(self.package)?,
        ))
    }
}

/// Contextual semantic view of one slide master.
pub struct SlideMaster<'a> {
    package: &'a OpcPackage,
    part: SlideMasterPart<'a>,
}

impl<'a> SlideMaster<'a> {
    /// Construct a master view from its package and low-level part view.
    pub fn new(package: &'a OpcPackage, part: SlideMasterPart<'a>) -> Self {
        Self { package, part }
    }

    /// Borrow the underlying master part view.
    #[inline]
    pub fn part(&self) -> &SlideMasterPart<'a> {
        &self.part
    }

    /// Return the producer name, falling back to the OPC part name.
    pub fn name(&self) -> Result<String> {
        self.part.name()
    }

    /// Whether PowerPoint should preserve the master after editing.
    pub fn is_preserved(&self) -> Result<bool> {
        self.part.is_preserved()
    }

    /// Build the bounded borrowed scene for the master.
    pub fn shapes(&self) -> Result<Scene<'a>> {
        self.part.shapes()
    }

    /// Read the theme reached from this slide master.
    pub fn theme(&self) -> Result<Option<crate::shape::theme::ThemeSummary>> {
        self.part.theme(self.package)
    }

    /// Resolve all layouts reachable from this master.
    pub fn layouts(&self) -> Result<Vec<SlideLayout<'a>>> {
        Ok(self
            .part
            .layouts(self.package)?
            .into_iter()
            .map(|part| SlideLayout::new(self.package, part))
            .collect())
    }

    /// Compatibility spelling used by existing contextual callers.
    pub fn slide_layouts(&self) -> Result<Vec<SlideLayout<'a>>> {
        self.layouts()
    }
}
