//! Typed PresentationML package facade and semantic package operations.

use litchi_opc::OpcPackage;
use litchi_opc::packuri::PackURI;

use crate::custom::{Host, Props};
use crate::parts::PresentationPart;
use crate::presentation::Presentation;
use crate::writer::MutablePresentation;
use crate::{Error, Result};

/// Main entry point for PresentationML package ownership.
pub struct Package {
    pub(crate) opc: OpcPackage,
    pub(crate) mutable_pres: Option<MutablePresentation>,
    #[cfg(feature = "encryption")]
    pub(crate) encryption: litchi_ooxml_common::package_encryption::PackageEncryption,
    #[cfg(feature = "automatic-fonts")]
    pub(crate) font_embedding_dirty: bool,
}

impl Package {
    /// Read the package's inert custom document properties.
    ///
    /// A package without a custom-properties part returns an empty collection.
    /// The shared OOXML owner validates the complete package-level relationship
    /// graph and preserves its bounded value semantics. See MS-OI29500 3.11,
    /// "Reserved Custom File Properties".
    pub fn custom_props(&self) -> Result<Props> {
        Ok(Props::read_for(&self.opc, Host::PowerPoint)?)
    }

    /// Replace the package's inert custom document properties transactionally.
    ///
    /// Empty properties remove the package-level relationship and target part.
    /// The shared OOXML owner validates and stages the graph before publication,
    /// preserving no-op and signature-invalidation behavior. See MS-OI29500
    /// 3.11, "Reserved Custom File Properties".
    pub fn put_custom_props(&mut self, props: Props) -> Result<()> {
        self.edit_typed(move |opc| Ok(props.write_for(opc, Host::PowerPoint)?))
    }

    /// Remove every inert custom document property transactionally.
    ///
    /// This is idempotent. See MS-OI29500 3.11, "Reserved Custom File
    /// Properties".
    pub fn remove_custom_props(&mut self) -> Result<()> {
        self.put_custom_props(Props::new())
    }

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
        self.ensure_plain_mutation("presentation_mut")?;
        self.mutable_pres.as_mut().ok_or(Error::UnsafeEdit {
            operation: "presentation_mut",
            reason: "the lossless facade cannot hydrate an opened package into the mutable writer",
        })
    }

    /// Read the presentation-owned embedded-font collection.
    pub fn fonts(&self) -> Result<Option<crate::font::Fonts>> {
        self.ensure_graph_current("fonts")?;
        crate::font::load(&self.opc)
    }

    /// Replace the complete presentation-owned embedded-font collection.
    ///
    /// This is the explicit, inert font-resource API. It remains independent
    /// from the optional system-font discovery policy used at managed save.
    pub fn put_fonts(&mut self, fonts: crate::font::Fonts) -> Result<bool> {
        self.edit_typed(move |opc| crate::font::put(opc, fonts))
    }

    /// Remove all presentation-owned embedded fonts and orphaned resources.
    pub fn remove_fonts(&mut self) -> Result<Option<crate::font::Fonts>> {
        self.edit_typed(crate::font::remove)
    }

    /// Whether a mutable model is currently pending managed publication.
    pub fn is_modified(&self) -> bool {
        self.mutable_pres
            .as_ref()
            .is_some_and(MutablePresentation::is_modified)
    }

    /// Borrow the underlying clear OPC graph when it is current.
    ///
    /// For a package with retained encryption provenance, this is an explicit
    /// declassification boundary: the borrowed graph is plaintext. Merely
    /// borrowing it does not discard the retained output policy.
    pub fn opc(&self) -> Result<&OpcPackage> {
        self.ensure_graph_current("opc")?;
        Ok(&self.opc)
    }

    /// Run a read-only operation against the current clear OPC graph.
    ///
    /// For encrypted provenance this explicitly exposes plaintext OPC content;
    /// retained output policy remains active after the closure returns.
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

    /// Read the source-bound Designer tags owned by one stable slide ID.
    pub fn slide_designer_tags<'a>(
        &self,
        slide: impl Into<crate::slide::Key<'a>>,
    ) -> Result<crate::presentation_properties::metadata::designer_tags::Snapshot> {
        self.slide_designer_tags_with_limits(slide, crate::shape::designer::Limits::default())
    }

    /// Read slide-ID Designer tags under caller-supplied resource bounds.
    pub fn slide_designer_tags_with_limits<'a>(
        &self,
        slide: impl Into<crate::slide::Key<'a>>,
        limits: crate::shape::designer::Limits,
    ) -> Result<crate::presentation_properties::metadata::designer_tags::Snapshot> {
        self.ensure_graph_current("slide_designer_tags")?;
        let (_, slide_id) = self.resolve_slide_identity(slide.into())?;
        crate::presentation_properties::metadata::designer_tags::load_snapshot_with_limits(
            &self.opc, slide_id, limits,
        )
    }

    /// Create or replace Designer tags on one stable slide ID atomically.
    pub fn put_slide_designer_tags<'a>(
        &mut self,
        slide: impl Into<crate::slide::Key<'a>>,
        tags: crate::shape::designer::Tags,
    ) -> Result<crate::presentation_properties::metadata::designer_tags::Snapshot> {
        self.put_slide_designer_tags_with_limits(
            slide,
            tags,
            crate::shape::designer::Limits::default(),
        )
    }

    /// Create or replace slide-ID Designer tags under explicit bounds.
    pub fn put_slide_designer_tags_with_limits<'a>(
        &mut self,
        slide: impl Into<crate::slide::Key<'a>>,
        tags: crate::shape::designer::Tags,
        limits: crate::shape::designer::Limits,
    ) -> Result<crate::presentation_properties::metadata::designer_tags::Snapshot> {
        self.ensure_graph_current("put_slide_designer_tags")?;
        let (_, slide_id) = self.resolve_slide_identity(slide.into())?;
        let mut edit =
            crate::presentation_properties::metadata::designer_tags::load_snapshot_with_limits(
                &self.opc, slide_id, limits,
            )?
            .edit()?;
        edit.set(tags)?;
        let commit = edit.commit()?;
        let changed = commit.is_changed();
        self.publish_designer(changed, move |opc| commit.apply(opc))
    }

    /// Remove Designer tags from one stable slide ID atomically.
    pub fn remove_slide_designer_tags<'a>(
        &mut self,
        slide: impl Into<crate::slide::Key<'a>>,
    ) -> Result<crate::presentation_properties::metadata::designer_tags::Snapshot> {
        self.remove_slide_designer_tags_with_limits(
            slide,
            crate::shape::designer::Limits::default(),
        )
    }

    /// Remove slide-ID Designer tags under explicit resource bounds.
    pub fn remove_slide_designer_tags_with_limits<'a>(
        &mut self,
        slide: impl Into<crate::slide::Key<'a>>,
        limits: crate::shape::designer::Limits,
    ) -> Result<crate::presentation_properties::metadata::designer_tags::Snapshot> {
        self.ensure_graph_current("remove_slide_designer_tags")?;
        let (_, slide_id) = self.resolve_slide_identity(slide.into())?;
        let mut edit =
            crate::presentation_properties::metadata::designer_tags::load_snapshot_with_limits(
                &self.opc, slide_id, limits,
            )?
            .edit()?;
        edit.remove();
        let commit = edit.commit()?;
        let changed = commit.is_changed();
        self.publish_designer(changed, move |opc| commit.apply(opc))
    }

    /// Read the typed document-level math defaults from presentation properties.
    pub fn math_properties(
        &self,
    ) -> Result<Option<crate::presentation_properties::math::Properties>> {
        self.ensure_graph_current("math_properties")?;
        crate::presentation_properties::load_math_from_package(&self.opc)
    }

    /// Replace the document-level math defaults transactionally.
    pub fn put_math_properties(
        &mut self,
        value: crate::presentation_properties::math::Properties,
    ) -> Result<Option<crate::presentation_properties::math::Properties>> {
        self.ensure_graph_current("put_math_properties")?;
        self.edit_typed(move |opc| crate::presentation_properties::put_math_to_package(opc, value))
    }

    /// Remove the document-level math defaults transactionally.
    pub fn remove_math_properties(
        &mut self,
    ) -> Result<Option<crate::presentation_properties::math::Properties>> {
        self.ensure_graph_current("remove_math_properties")?;
        self.edit_typed(crate::presentation_properties::remove_math_from_package)
    }

    /// Read the lossless zoom owner of one slide.
    pub fn zooms<'a>(
        &self,
        slide: impl Into<crate::slide::Key<'a>>,
    ) -> Result<crate::shape::zoom::Owner> {
        self.ensure_graph_current("zooms")?;
        let slide_name = self.resolve_slide(slide.into())?;
        crate::shape::zoom::load(&self.opc, &slide_name)
    }

    /// Replace one slide's zoom owner transactionally.
    pub fn put_zooms<'a>(
        &mut self,
        slide: impl Into<crate::slide::Key<'a>>,
        owner: crate::shape::zoom::Owner,
    ) -> Result<Option<crate::shape::zoom::Owner>> {
        self.ensure_graph_current("put_zooms")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::shape::zoom::store(opc, &slide_name, owner))
    }

    /// Remove all zoom metadata from one slide while retaining its other XML.
    pub fn remove_zooms<'a>(
        &mut self,
        slide: impl Into<crate::slide::Key<'a>>,
    ) -> Result<Option<crate::shape::zoom::Owner>> {
        self.ensure_graph_current("remove_zooms")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::shape::zoom::remove(opc, &slide_name))
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

    /// Read the optional classification outcome attached to one semantic
    /// shape. The shape selector uses its exact producer name by default and
    /// retains checked source-order indices for repair workflows.
    pub fn shape_classification<'s, 'k>(
        &self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::shape::classification::Snapshot>> {
        self.ensure_graph_current("shape_classification")?;
        let slide_name = self.resolve_slide(slide.into())?;
        crate::shape::classification::load(&self.opc, &slide_name, shape)
    }

    /// Create or replace one semantic shape's typed classification outcome
    /// transactionally while retaining unrelated extension markup.
    pub fn put_shape_classification<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        outcome: crate::shape::classification::Outcome,
    ) -> Result<Option<crate::shape::classification::Snapshot>> {
        self.ensure_graph_current("put_shape_classification")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| {
            crate::shape::classification::put(opc, &slide_name, shape, outcome)
        })
    }

    /// Remove one semantic shape's typed classification element
    /// transactionally. Unknown extension entries remain intact.
    pub fn remove_shape_classification<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::shape::classification::Snapshot>> {
        self.ensure_graph_current("remove_shape_classification")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::shape::classification::remove(opc, &slide_name, shape))
    }

    /// Read the optional `p15:designElem` value attached to one semantic
    /// shape. The selector is name-first, with checked source-order indices
    /// retained for repair workflows.
    pub fn shape_design_element<'s, 'k>(
        &self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::shape::designer::Snapshot>> {
        self.ensure_graph_current("shape_design_element")?;
        let slide_name = self.resolve_slide(slide.into())?;
        crate::shape::designer::load(&self.opc, &slide_name, shape)
    }

    /// Create or replace one shape's typed `p15:designElem` boolean
    /// transactionally while retaining unrelated extension entries.
    pub fn put_shape_design_element<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        value: bool,
    ) -> Result<Option<crate::shape::designer::Snapshot>> {
        self.ensure_graph_current("put_shape_design_element")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::shape::designer::put(opc, &slide_name, shape, value))
    }

    /// Remove one shape's typed `p15:designElem` value transactionally.
    pub fn remove_shape_design_element<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::shape::designer::Snapshot>> {
        self.ensure_graph_current("remove_shape_design_element")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::shape::designer::remove(opc, &slide_name, shape))
    }

    /// Read one shape's source-bound PowerPoint 2020 Designer properties.
    pub fn shape_designer_properties<'s, 'k>(
        &self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<crate::shape::designer::PropertiesSnapshot> {
        self.shape_designer_properties_with_limits(
            slide,
            shape,
            crate::shape::designer::Limits::default(),
        )
    }

    /// Read shape Designer properties under caller-supplied resource bounds.
    pub fn shape_designer_properties_with_limits<'s, 'k>(
        &self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        limits: crate::shape::designer::Limits,
    ) -> Result<crate::shape::designer::PropertiesSnapshot> {
        self.ensure_graph_current("shape_designer_properties")?;
        let (slide_name, _) = self.resolve_slide_identity(slide.into())?;
        crate::shape::designer::load_properties_with_limits(&self.opc, &slide_name, shape, limits)
    }

    /// Create or replace one shape's Designer drawing properties atomically.
    pub fn put_shape_designer_properties<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        properties: crate::shape::designer::DrawingProperties,
    ) -> Result<crate::shape::designer::PropertiesSnapshot> {
        self.ensure_graph_current("put_shape_designer_properties")?;
        let (slide_name, _) = self.resolve_slide_identity(slide.into())?;
        let shape = shape.into();
        let before = crate::shape::designer::load_properties(&self.opc, &slide_name, shape)?;
        let changed = before.properties() != Some(&properties);
        self.publish_designer(changed, move |opc| {
            crate::shape::designer::put_properties(opc, &slide_name, shape, properties)
        })
    }

    /// Create or replace shape Designer properties under explicit bounds.
    pub fn put_shape_designer_properties_with_limits<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        properties: crate::shape::designer::DrawingProperties,
        limits: crate::shape::designer::Limits,
    ) -> Result<crate::shape::designer::PropertiesSnapshot> {
        self.ensure_graph_current("put_shape_designer_properties")?;
        let (slide_name, _) = self.resolve_slide_identity(slide.into())?;
        let mut edit = crate::shape::designer::load_properties_with_limits(
            &self.opc,
            &slide_name,
            shape,
            limits,
        )?
        .edit();
        edit.set(properties)?;
        let commit = edit.commit()?;
        let changed = !commit.is_noop();
        self.publish_designer(changed, move |opc| {
            crate::shape::designer::apply_properties(opc, commit)
        })
    }

    /// Remove one shape's Designer drawing properties atomically.
    pub fn remove_shape_designer_properties<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<crate::shape::designer::PropertiesSnapshot> {
        self.ensure_graph_current("remove_shape_designer_properties")?;
        let (slide_name, _) = self.resolve_slide_identity(slide.into())?;
        let shape = shape.into();
        let before = crate::shape::designer::load_properties(&self.opc, &slide_name, shape)?;
        let changed = before.is_present();
        self.publish_designer(changed, move |opc| {
            crate::shape::designer::remove_properties(opc, &slide_name, shape)
        })
    }

    /// Remove shape Designer properties under explicit resource bounds.
    pub fn remove_shape_designer_properties_with_limits<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        limits: crate::shape::designer::Limits,
    ) -> Result<crate::shape::designer::PropertiesSnapshot> {
        self.ensure_graph_current("remove_shape_designer_properties")?;
        let (slide_name, _) = self.resolve_slide_identity(slide.into())?;
        let mut edit = crate::shape::designer::load_properties_with_limits(
            &self.opc,
            &slide_name,
            shape,
            limits,
        )?
        .edit();
        edit.remove();
        let commit = edit.commit()?;
        let changed = !commit.is_noop();
        self.publish_designer(changed, move |opc| {
            crate::shape::designer::apply_properties(opc, commit)
        })
    }

    /// Read all contextual 3D-model owners on one slide in source order.
    pub fn model3ds<'a>(
        &self,
        slide: impl Into<crate::slide::Key<'a>>,
    ) -> Result<Vec<crate::model3d::Model>> {
        self.ensure_graph_current("model3ds")?;
        let slide_name = self.resolve_slide(slide.into())?;
        crate::model3d::package::load_all(&self.opc, &slide_name)
    }

    /// Read the model3d owner attached to one semantic shape, if present.
    pub fn model3d<'s, 'k>(
        &self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::model3d::Model>> {
        self.ensure_graph_current("model3d")?;
        let slide_name = self.resolve_slide(slide.into())?;
        crate::model3d::package::load(&self.opc, &slide_name, shape.into())
    }

    /// Replace one existing model3d owner transactionally.
    pub fn put_model3d<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        model: crate::model3d::Model,
    ) -> Result<Option<crate::model3d::Model>> {
        self.ensure_graph_current("put_model3d")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| {
            crate::model3d::package::put(opc, &slide_name, shape.into(), model)
        })
    }

    /// Remove one model3d owner and collect unreachable binary resources.
    pub fn remove_model3d<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::model3d::Model>> {
        self.ensure_graph_current("remove_model3d")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::model3d::package::remove(opc, &slide_name, shape.into()))
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

    /// Run one transactional low-level clear-OPC edit.
    ///
    /// The closure receives the current graph only after pending authoring
    /// state has been published. Any error rolls the graph back to its exact
    /// pre-edit snapshot; successful edits commit the candidate in place.
    /// Callers that need a typed, semantic operation should prefer the
    /// contextual methods on this facade. For retained encryption provenance,
    /// choosing this raw escape hatch explicitly declassifies the package. A
    /// successful edit clears provenance; a failed edit preserves it.
    pub fn edit_opc<T>(
        &mut self,
        operation: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        let value = self.edit_raw(|opc| {
            let signed_source = opc
                .is_signed()
                .then(|| litchi_opc::PackageWriter::to_bytes(opc))
                .transpose()?;
            let value = operation(opc)?;
            if let Some(source) = signed_source {
                let candidate = litchi_opc::PackageWriter::to_bytes(opc)?;
                if candidate != source {
                    opc.unsign();
                }
            }
            Ok(value)
        })?;
        // A raw graph edit cannot be reflected into the lossless mutable
        // writer. Retire that facade after publication so later authoring
        // cannot overwrite the committed OPC graph with stale state.
        self.mutable_pres = None;
        #[cfg(feature = "encryption")]
        {
            self.encryption = litchi_ooxml_common::package_encryption::PackageEncryption::plain();
        }
        Ok(value)
    }

    pub(crate) fn ensure_plain_mutation(&self, operation: &'static str) -> Result<()> {
        let _ = operation;
        #[cfg(feature = "encryption")]
        self.encryption
            .ordinary_output()
            .map_err(|source| Error::EncryptionPolicy { operation, source })?;
        Ok(())
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

    fn publish_designer<T>(
        &mut self,
        changed: bool,
        operation: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        let value = self.edit_typed(operation)?;
        if changed {
            self.mutable_pres = None;
        }
        Ok(value)
    }

    fn resolve_slide_identity(&self, key: crate::slide::Key<'_>) -> Result<(PackURI, u32)> {
        let presentation = self.presentation()?;
        let references = presentation.slide_references()?;
        match key {
            crate::slide::Key::Index(index) => {
                let reference = references.get(index).ok_or(Error::SlideIndexOutOfBounds {
                    index,
                    len: references.len(),
                })?;
                let slide = presentation
                    .slide(index)?
                    .ok_or(Error::SlideIndexOutOfBounds {
                        index,
                        len: references.len(),
                    })?;
                Ok((slide.part().part().partname().clone(), reference.id()))
            },
            crate::slide::Key::Name(name) => {
                let slides = presentation.slides()?;
                let mut selected = None;
                let mut matches = 0usize;
                for (reference, slide) in references.iter().zip(slides) {
                    if slide.name()? == name {
                        matches = matches.saturating_add(1);
                        selected = Some((slide.part().part().partname().clone(), reference.id()));
                    }
                }
                match (matches, selected) {
                    (0, _) => Err(Error::SlideNameNotFound(name.to_owned())),
                    (1, Some(value)) => Ok(value),
                    _ => Err(Error::AmbiguousSlideName {
                        name: name.to_owned(),
                        matches,
                    }),
                }
            },
        }
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
