//! Lossless-bounded PPTX package facade.

use std::io::Read;
use std::path::Path;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::packuri::PackURI;
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackageWriter};

use crate::parts::PresentationPart;
use crate::presentation::Presentation;
use crate::resources;
use crate::writer::MutablePresentation;
use crate::{Error, Result};

const PRESENTATION_PART: &str = "/ppt/presentation.xml";
const MASTER_PART: &str = "/ppt/slideMasters/slideMaster1.xml";
const THEME_PART: &str = "/ppt/theme/theme1.xml";
const VIEW_PROPERTIES_PART: &str = "/ppt/viewProps.xml";
const PRESENTATION_PROPERTIES_PART: &str = "/ppt/presProps.xml";
const TABLE_STYLES_PART: &str = "/ppt/tableStyles.xml";
const NOTES_THEME_PART: &str = "/ppt/theme/theme2.xml";
const NOTES_MASTER_PART: &str = "/ppt/notesMasters/notesMaster1.xml";
const NOTES_MASTER_RELATIONSHIP_ID: &str = "rIdNotesMaster";
const CORE_PROPERTIES_PART: &str = "/docProps/core.xml";
const EXTENDED_PROPERTIES_PART: &str = "/docProps/app.xml";

const STALE_PRESENTATION_GRAPH_REASON: &str = "the mutable presentation model has unflushed changes; save and reopen before reading the canonical package graph";

/// Main entry point for PresentationML package ownership.
pub struct Package {
    opc: OpcPackage,
    mutable_pres: Option<MutablePresentation>,
}

impl Package {
    /// Create a minimal valid, mutable PresentationML package.
    pub fn new() -> Result<Self> {
        let mut package = OpcPackage::new();
        let presentation_name = pack_uri(PRESENTATION_PART)?;
        let master_name = pack_uri(MASTER_PART)?;
        let theme_name = pack_uri(THEME_PART)?;
        let view_properties_name = pack_uri(VIEW_PROPERTIES_PART)?;
        let presentation_properties_name = pack_uri(PRESENTATION_PROPERTIES_PART)?;
        let table_styles_name = pack_uri(TABLE_STYLES_PART)?;
        let notes_theme_name = pack_uri(NOTES_THEME_PART)?;
        let notes_master_name = pack_uri(NOTES_MASTER_PART)?;
        let core_properties_name = pack_uri(CORE_PROPERTIES_PART)?;
        let extended_properties_name = pack_uri(EXTENDED_PROPERTIES_PART)?;

        let presentation_xml = resources::PRESENTATION.replacen(
            "</p:sldMasterIdLst>",
            &format!(
                "</p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id=\"{NOTES_MASTER_RELATIONSHIP_ID}\"/></p:notesMasterIdLst>"
            ),
            1,
        );
        let mut presentation = BlobPart::new(
            presentation_name.clone(),
            ct::PML_PRESENTATION_MAIN.to_string(),
            presentation_xml.into_bytes(),
        );
        presentation.relate_to("slideMasters/slideMaster1.xml", rt::SLIDE_MASTER);
        presentation.relate_to("viewProps.xml", rt::VIEW_PROPS);
        presentation.relate_to("presProps.xml", rt::PRES_PROPS);
        // Keep the conventional numeric relationship IDs available to the
        // presentation writer. The optional catalog has a semantic ID and
        // must not shift slide relationship IDs in newly authored packages.
        presentation.rels_mut().add_relationship(
            rt::TABLE_STYLES.to_string(),
            "tableStyles.xml".to_string(),
            "rIdTableStyles".to_string(),
            false,
        );
        presentation.rels_mut().add_relationship(
            rt::NOTES_MASTER.to_string(),
            "notesMasters/notesMaster1.xml".to_string(),
            NOTES_MASTER_RELATIONSHIP_ID.to_string(),
            false,
        );

        let mut master = BlobPart::new(
            master_name,
            ct::PML_SLIDE_MASTER.to_string(),
            resources::SLIDE_MASTER.as_bytes().to_vec(),
        );
        for index in 1..=resources::SLIDE_LAYOUTS.len() {
            master.relate_to(
                &format!("../slideLayouts/slideLayout{index}.xml"),
                rt::SLIDE_LAYOUT,
            );
        }
        master.relate_to("../theme/theme1.xml", rt::THEME);

        let theme = BlobPart::new(
            theme_name,
            ct::OFC_THEME.to_string(),
            resources::THEME.as_bytes().to_vec(),
        );
        let notes_theme = BlobPart::new(
            notes_theme_name,
            ct::OFC_THEME.to_string(),
            resources::THEME.as_bytes().to_vec(),
        );
        let mut notes_master = BlobPart::new(
            notes_master_name,
            ct::PML_NOTES_MASTER.to_string(),
            crate::notes::master_xml().as_bytes().to_vec(),
        );
        notes_master.relate_to("../theme/theme2.xml", rt::THEME);

        let view_properties = BlobPart::new(
            view_properties_name,
            ct::PML_VIEW_PROPS.to_string(),
            resources::VIEW_PROPERTIES.as_bytes().to_vec(),
        );
        let presentation_properties = BlobPart::new(
            presentation_properties_name,
            ct::PML_PRES_PROPS.to_string(),
            resources::PRESENTATION_PROPERTIES.as_bytes().to_vec(),
        );
        let table_styles = BlobPart::new(
            table_styles_name,
            ct::PML_TABLE_STYLES.to_string(),
            crate::table::style::default_xml().as_bytes().to_vec(),
        );
        let core_properties = BlobPart::new(
            core_properties_name,
            ct::OPC_CORE_PROPERTIES.to_string(),
            resources::CORE_PROPERTIES.as_bytes().to_vec(),
        );
        let extended_properties = BlobPart::new(
            extended_properties_name,
            ct::OFC_EXTENDED_PROPERTIES.to_string(),
            resources::EXTENDED_PROPERTIES.as_bytes().to_vec(),
        );

        package.relate_to("ppt/presentation.xml", rt::OFFICE_DOCUMENT);
        package.relate_to("docProps/core.xml", rt::CORE_PROPERTIES);
        package.relate_to("docProps/app.xml", rt::EXTENDED_PROPERTIES);
        package.add_part(Box::new(presentation));
        package.add_part(Box::new(master));
        for (index, xml) in resources::SLIDE_LAYOUTS.iter().enumerate() {
            let layout_name = pack_uri(&format!("/ppt/slideLayouts/slideLayout{}.xml", index + 1))?;
            let mut layout = BlobPart::new(
                layout_name,
                ct::PML_SLIDE_LAYOUT.to_string(),
                xml.as_bytes().to_vec(),
            );
            layout.relate_to("../slideMasters/slideMaster1.xml", rt::SLIDE_MASTER);
            package.add_part(Box::new(layout));
        }
        package.add_part(Box::new(theme));
        package.add_part(Box::new(notes_theme));
        package.add_part(Box::new(notes_master));
        package.add_part(Box::new(view_properties));
        package.add_part(Box::new(presentation_properties));
        package.add_part(Box::new(table_styles));
        package.add_part(Box::new(core_properties));
        package.add_part(Box::new(extended_properties));

        Ok(Self {
            opc: package,
            mutable_pres: Some(MutablePresentation::new()),
        })
    }

    /// Open a PPTX from a filesystem path.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::from_opc_package(OpcPackage::open(path)?)
    }

    /// Parse a PPTX from a reader.
    pub fn from_reader<R: Read>(reader: R) -> Result<Self> {
        Self::from_opc_package(OpcPackage::from_reader(reader)?)
    }

    /// Parse a PPTX from an owned ZIP buffer.
    pub fn from_vec(bytes: Vec<u8>) -> Result<Self> {
        Self::from_opc_package(OpcPackage::from_vec(bytes)?)
    }

    /// Parse a PPTX from a borrowed ZIP buffer.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_opc_package(OpcPackage::from_bytes(bytes)?)
    }

    /// Adopt an already parsed OPC graph without hydrating a mutable model.
    pub fn from_opc_package(opc: OpcPackage) -> Result<Self> {
        let main = opc.main_document_part()?;
        PresentationPart::from_part(main)?;
        Ok(Self {
            opc,
            mutable_pres: None,
        })
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

    /// Save the package atomically through the OPC writer.
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        self.flush_presentation()?;
        PackageWriter::write(path, &self.opc)?;
        Ok(())
    }

    /// Serialize the package into a new ZIP buffer.
    pub fn to_bytes(&mut self) -> Result<Vec<u8>> {
        self.flush_presentation()?;
        Ok(PackageWriter::to_bytes(&self.opc)?)
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

    fn edit_typed<T>(&mut self, operation: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        self.flush_presentation()?;
        let before = self.opc.clone();
        match operation(&mut self.opc) {
            Ok(value) => Ok(value),
            Err(error) => {
                self.opc = before;
                Err(error)
            },
        }
    }

    fn ensure_graph_current(&self, operation: &'static str) -> Result<()> {
        if self.is_modified() {
            return Err(Error::UnsafeEdit {
                operation,
                reason: STALE_PRESENTATION_GRAPH_REASON,
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

    fn flush_presentation(&mut self) -> Result<()> {
        if !self.is_modified() {
            return Ok(());
        }
        let Some(presentation) = self.mutable_pres.as_ref().cloned() else {
            return Err(Error::UnsafeEdit {
                operation: "save",
                reason: "the lossless facade cannot publish an opened package's mutable graph",
            });
        };
        let before = self.opc.clone();
        match self.materialize_presentation(&presentation) {
            Ok(()) => {
                if let Some(presentation) = self.mutable_pres.as_mut() {
                    presentation.mark_clean();
                }
                Ok(())
            },
            Err(error) => {
                self.opc = before;
                Err(error)
            },
        }
    }

    fn materialize_presentation(&mut self, presentation: &MutablePresentation) -> Result<()> {
        let presentation_name = pack_uri(PRESENTATION_PART)?;
        let old_slide_relationships: Vec<String> = {
            let part = self.opc.get_part(&presentation_name)?;
            part.rels()
                .iter()
                .filter(|relationship| {
                    crate::parts::is_relationship_type(relationship.reltype(), rt::SLIDE, "slide")
                })
                .map(|relationship| relationship.r_id().to_string())
                .collect()
        };
        {
            let part = self.opc.get_part_mut(&presentation_name)?;
            for relationship_id in old_slide_relationships {
                part.rels_mut().remove(&relationship_id);
            }
        }

        let old_slides: Vec<_> = self
            .opc
            .iter_parts()
            .filter(|part| part.partname().as_str().starts_with("/ppt/slides/"))
            .map(|part| part.partname().clone())
            .collect();
        for part in old_slides {
            self.opc.remove_part(&part);
        }
        let old_notes_slides: Vec<_> = self
            .opc
            .iter_parts()
            .filter(|part| part.partname().as_str().starts_with("/ppt/notesSlides/"))
            .map(|part| part.partname().clone())
            .collect();
        for part in old_notes_slides {
            self.opc.remove_part(&part);
        }

        let mut relationship_ids = Vec::with_capacity(presentation.slide_count());
        for (index, slide) in presentation.slides().iter().enumerate() {
            let target = format!("slides/slide{}.xml", index.saturating_add(1));
            let relationship_id = self
                .opc
                .get_part_mut(&presentation_name)?
                .relate_to(&target, rt::SLIDE);
            relationship_ids.push(relationship_id);

            let slide_name = pack_uri(&format!("/ppt/{target}"))?;
            let mut slide_part = BlobPart::new(
                slide_name,
                ct::PML_SLIDE.to_string(),
                slide.generate_slide_xml()?.into_bytes(),
            );
            slide_part.relate_to("../slideLayouts/slideLayout1.xml", rt::SLIDE_LAYOUT);
            if slide.has_notes() {
                slide_part.relate_to(
                    &format!("../notesSlides/notesSlide{}.xml", index + 1),
                    rt::NOTES_SLIDE,
                );
            }
            self.opc.add_part(Box::new(slide_part));

            if let Some(text) = slide.notes() {
                let notes_name =
                    pack_uri(&format!("/ppt/notesSlides/notesSlide{}.xml", index + 1))?;
                let mut notes_part = BlobPart::new(
                    notes_name,
                    ct::PML_NOTES_SLIDE.to_string(),
                    crate::notes::write_text(text)?,
                );
                notes_part.relate_to(&format!("../slides/slide{}.xml", index + 1), rt::SLIDE);
                notes_part.relate_to("../notesMasters/notesMaster1.xml", rt::NOTES_MASTER);
                self.opc.add_part(Box::new(notes_part));
            }
        }

        let xml = presentation.generate_presentation_xml_with(&relationship_ids)?;
        self.opc
            .get_part_mut(&presentation_name)?
            .set_blob(xml.into_bytes());
        Ok(())
    }
}

fn pack_uri(value: &str) -> Result<PackURI> {
    PackURI::new(value).map_err(Error::Uri)
}

#[cfg(test)]
mod tests {
    use super::Package;

    #[test]
    fn new_writer_round_trips_the_bounded_slide_graph() {
        let mut package = Package::new().expect("new package");
        {
            let presentation = package.presentation_mut().expect("mutable presentation");
            let slide = presentation.add_slide().expect("slide");
            slide.set_title("Canonical owner");
            slide.add_text_box("Hello & goodbye", 914_400, 914_400, 2_743_200, 914_400);
            presentation.set_widescreen_slide_size();
        }

        let bytes = package.to_bytes().expect("serialize package");
        let reopened = Package::from_bytes(&bytes).expect("reopen package");
        let presentation = reopened.presentation().expect("presentation");
        assert_eq!(presentation.slide_count().expect("slide count"), 1);
        assert_eq!(
            presentation.slide_size().expect("slide size"),
            (9_144_000, 5_143_500)
        );
        let slide = presentation.slide(0).expect("slide lookup").expect("slide");
        assert_eq!(slide.name().expect("slide name"), "Slide 256");
        assert!(
            slide
                .text()
                .expect("slide text")
                .contains("Hello & goodbye")
        );
        assert_eq!(slide.shape_count().expect("shape count"), 2);
        assert_eq!(presentation.slide_masters().expect("masters").len(), 1);
        assert_eq!(presentation.slide_layouts().expect("layouts").len(), 11);
    }

    #[test]
    fn opened_package_refuses_unsafe_mutable_hydration() {
        let mut package = Package::new().expect("new package");
        let bytes = package.to_bytes().expect("serialize package");
        let mut opened = Package::from_bytes(&bytes).expect("reopen package");
        assert!(opened.presentation_mut().is_err());
    }
}
