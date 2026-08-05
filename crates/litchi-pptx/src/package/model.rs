//! Typed PresentationML package facade and semantic package operations.

use litchi_opc::OpcPackage;
use litchi_opc::packuri::PackURI;

use crate::parts::PresentationPart;
use crate::presentation::Presentation;
use crate::writer::MutablePresentation;
use crate::{Error, Result};

/// Main entry point for PresentationML package ownership.
pub struct Package {
    pub(super) opc: OpcPackage,
    pub(super) mutable_pres: Option<MutablePresentation>,
}

impl Package {
    /// Borrow the canonical presentation graph when no mutable state is stale.
    pub fn presentation(&self) -> Result<Presentation<'_>> {
        self.ensure_graph_current("presentation")?;
        Ok(Presentation::new(
            PresentationPart::from_package(&self.opc)?,
            &self.opc,
        ))
    }

    /// Borrow the mutable presentation model for a newly authored package.
    pub fn presentation_mut(&mut self) -> Result<&mut MutablePresentation> {
        self.mutable_pres.as_mut().ok_or(Error::UnsafeEdit {
            operation: "presentation_mut",
            reason: "the lossless facade cannot hydrate an opened package into the mutable writer",
        })
    }

    /// Whether a mutable model is currently pending managed publication.
    pub fn is_modified(&self) -> bool {
        self.mutable_pres
            .as_ref()
            .is_some_and(MutablePresentation::is_modified)
    }

    /// Borrow the underlying OPC graph when it is current.
    pub fn opc(&self) -> Result<&OpcPackage> {
        self.ensure_graph_current("opc")?;
        Ok(&self.opc)
    }

    /// Run a read-only operation against the current OPC graph.
    pub fn with_opc<T>(&self, operation: impl FnOnce(&OpcPackage) -> Result<T>) -> Result<T> {
        self.ensure_graph_current("with_opc")?;
        operation(&self.opc)
    }

    /// Read one slide's direct, inert programmable-tag list.
    ///
    /// Names are the ordinary selector and zero-based presentation positions
    /// remain available for ordered repair workflows. The returned list owns
    /// its bounded strings, so the read does not borrow the package graph.
    pub fn tags<'a>(
        &self,
        slide: impl Into<crate::slide::Key<'a>>,
    ) -> Result<Option<crate::tag::List>> {
        self.ensure_graph_current("tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        Ok(crate::tag::load(&self.opc, &slide_name)?.map(crate::tag::Source::into_list))
    }

    /// Create or replace one slide's direct programmable-tag list.
    ///
    /// The list is moved into the package transaction and the staged owner
    /// relationship and part are published together.
    pub fn put_tags<'a>(
        &mut self,
        slide: impl Into<crate::slide::Key<'a>>,
        list: crate::tag::List,
    ) -> Result<Option<crate::tag::List>> {
        self.ensure_graph_current("put_tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::tag::put(opc, &slide_name, list))
    }

    /// Remove one slide's direct programmable-tag list.
    ///
    /// Removal is idempotent and only collects an orphaned tag part after the
    /// package-wide inbound-edge check succeeds.
    pub fn remove_tags<'a>(
        &mut self,
        slide: impl Into<crate::slide::Key<'a>>,
    ) -> Result<Option<crate::tag::List>> {
        self.ensure_graph_current("remove_tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::tag::remove(opc, &slide_name))
    }

    /// Read one semantic shape's optional programmable-tag list.
    pub fn shape_tags<'s, 'k>(
        &self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::tag::List>> {
        self.ensure_graph_current("shape_tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        Ok(crate::tag::shape::load(&self.opc, &slide_name, shape)?
            .map(crate::tag::Source::into_list))
    }

    /// Create or replace one semantic shape's programmable-tag list.
    pub fn put_shape_tags<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        list: crate::tag::List,
    ) -> Result<Option<crate::tag::List>> {
        self.ensure_graph_current("put_shape_tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::tag::shape::put(opc, &slide_name, shape, list))
    }

    /// Remove one semantic shape's programmable-tag list.
    pub fn remove_shape_tags<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::tag::List>> {
        self.ensure_graph_current("remove_shape_tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::tag::shape::remove(opc, &slide_name, shape))
    }

    /// Load the presentation's optional DrawingML table-style catalog.
    pub fn styles(&self) -> Result<Option<crate::table::style::List>> {
        // Table styles are owned by their OPC part, independently of the
        // slide-authoring model. They remain a safe immutable read while a
        // newly authored slide is still pending publication.
        crate::table::style::load(&self.opc)
    }

    /// Create or replace the presentation's table-style catalog atomically.
    pub fn put_styles(&mut self, styles: crate::table::style::List) -> Result<bool> {
        self.edit_typed(move |opc| crate::table::style::put(opc, styles))
    }

    /// Remove the presentation's optional table-style catalog atomically.
    pub fn remove_styles(&mut self) -> Result<Option<crate::table::style::List>> {
        self.edit_typed(crate::table::style::remove)
    }

    /// Add a slide master and update the PresentationML relationship graph.
    pub fn add_slide_master(&mut self) -> Result<crate::master_layout::AuthoredSlideMaster> {
        self.edit_typed(crate::master_layout::add_slide_master)
    }

    /// Add a layout to an existing master and update both sides of the graph.
    pub fn add_slide_layout(
        &mut self,
        master_part_name: &PackURI,
        kind: crate::master_layout::SlideLayoutKind,
        name: &str,
        placeholders: &[crate::master_layout::PlaceholderSpec],
    ) -> Result<crate::master_layout::AuthoredSlideLayout> {
        self.edit_typed(|opc| {
            crate::master_layout::add_slide_layout(opc, master_part_name, kind, name, placeholders)
        })
    }

    /// Add or replace one master/layout placeholder shape.
    pub fn store_placeholder_shape(
        &mut self,
        part_name: &PackURI,
        spec: &crate::master_layout::PlaceholderSpec,
    ) -> Result<()> {
        self.edit_typed(|opc| crate::master_layout::store_placeholder_shape(opc, part_name, spec))
    }

    /// Remove an unreferenced layout and its owning relationship.
    pub fn remove_slide_layout(&mut self, layout_part_name: &PackURI) -> Result<()> {
        self.edit_typed(|opc| crate::master_layout::remove_slide_layout(opc, layout_part_name))
    }

    /// Validate every master/layout relationship reachable from the package.
    pub fn validate_master_layout_graph(&self) -> Result<()> {
        self.with_opc(crate::master_layout::validate_master_layout_graph)
    }

    /// Load all contextual slide-library synchronization metadata.
    pub fn load_slide_sync(
        &self,
    ) -> Result<Vec<crate::presentation_properties::metadata::slide_sync::Part>> {
        self.with_opc(crate::presentation_properties::metadata::slide_sync::load)
    }

    /// Attach one slide-library synchronization part transactionally.
    pub fn store_slide_sync(
        &mut self,
        value: &crate::presentation_properties::metadata::slide_sync::Part,
    ) -> Result<()> {
        self.edit_typed(|opc| {
            crate::presentation_properties::metadata::slide_sync::store(opc, value)
        })
    }

    /// Run one transactional low-level OPC edit.
    ///
    /// The closure receives the current graph only after pending authoring
    /// state has been published. Any error rolls the graph back to its exact
    /// pre-edit snapshot; successful edits commit the candidate in place.
    /// Callers that need a typed, semantic operation should prefer the
    /// contextual methods on this facade.
    pub fn edit_opc<T>(
        &mut self,
        operation: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        let value = self.edit_typed(operation)?;
        // A raw graph edit cannot be reflected into the lossless mutable
        // writer. Retire that facade after publication so later authoring
        // cannot overwrite the committed OPC graph with stale state.
        self.mutable_pres = None;
        Ok(value)
    }

    fn ensure_graph_current(&self, operation: &'static str) -> Result<()> {
        if self.is_modified() {
            return Err(Error::UnsafeEdit {
                operation,
                reason: super::codec::STALE_PRESENTATION_GRAPH_REASON,
            });
        }
        Ok(())
    }

    fn resolve_slide(&self, key: crate::slide::Key<'_>) -> Result<PackURI> {
        let presentation = self.presentation()?;
        match key {
            crate::slide::Key::Index(index) => {
                let length = presentation.slide_count()?;
                let slide = presentation
                    .slide(index)?
                    .ok_or(Error::SlideIndexOutOfBounds { index, len: length })?;
                Ok(slide.part().part().partname().clone())
            },
            crate::slide::Key::Name(name) => {
                let slide = presentation
                    .find_slide(name)?
                    .ok_or_else(|| Error::SlideNameNotFound(name.to_owned()))?;
                Ok(slide.part().part().partname().clone())
            },
        }
    }
}
