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

    /// Count semantic shapes in the slide's scene.
    pub fn shape_count(&self) -> Result<usize> {
        Ok(self.shapes()?.len())
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
