//! OPC package graph construction, publication, and serialization.

use std::io::Read;
use std::path::Path;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::packuri::PackURI;
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackageWriter};

use super::Package;
use crate::custom::{Host, Props};
use crate::parts::PresentationPart;
use crate::resources;
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

impl Package {
    /// Create a minimal valid canonical `PresentationML` package.
    ///
    /// Use [`crate::builder::Builder`] to configure a fresh presentation.
    ///
    /// ```
    /// let package = litchi_pptx::Package::new()?;
    /// assert_eq!(package.presentation()?.slide_count()?, 0);
    /// let commit = package.edit()?.commit()?;
    /// assert!(commit.patch().is_noop());
    /// let mut bytes = Vec::new();
    /// commit.into_package().write_to(&mut bytes)?;
    /// # Ok::<(), litchi_pptx::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new() -> Result<Self> {
        Self::new_authoring()
    }

    pub(crate) fn new_authoring() -> Result<Self> {
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
            crate::notes::master_xml(crate::notes::Limits::DEFAULT)?
                .as_bytes()
                .to_vec(),
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
        let part_bytes = Self::part_bytes_for(&package)?;

        let replay_index = crate::package::replay_index::ReplayIndex::from_package(&package)?;

        Ok(Self {
            opc: std::sync::Arc::new(package),
            part_bytes,
            content_revision: std::sync::Arc::new(std::sync::OnceLock::new()),
            part_digests: std::sync::Arc::new(std::sync::OnceLock::new()),
            replay_index: std::sync::Arc::new(replay_index),
            physical_source_provenance: false,
            semantic_limits: None,
            tolerance: crate::leniency::ToleranceReport::default(),
            #[cfg(feature = "encryption")]
            encryption: litchi_ooxml_common::package_encryption::PackageEncryption::plain(),
            #[cfg(feature = "encryption")]
            encryption_limits: litchi_crypto::ooxml::Limits::default(),
        })
    }

    /// Open a PPTX from a filesystem path.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::open_with_options(path, crate::OpenOptions::default())
    }

    /// Open a PPTX from a filesystem path with explicit ingress options.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::UnsupportedOpenPolicy`] for cache or execution
    /// policies, which only positional-source owners honor, and an error if
    /// the input exceeds the source limits, cannot be read or is malformed.
    pub fn open_with_options<P: AsRef<Path>>(path: P, options: crate::OpenOptions) -> Result<Self> {
        options.check_owning()?;
        Self::from_opc_package_with_provenance_and_semantic_limits(
            OpcPackage::open(path, options.source_limits)?,
            true,
            options.semantic_limits,
        )
    }

    /// Parse a PPTX from a reader.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_reader<R: Read>(reader: R) -> Result<Self> {
        Self::from_reader_with_options(reader, crate::OpenOptions::default())
    }

    /// Parse a PPTX from a reader with explicit ingress options.
    ///
    /// Reader ingestion is bounded by the selected source limits.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::UnsupportedOpenPolicy`] for cache or execution
    /// policies, and an error if the input exceeds the source limits, cannot
    /// be read or is malformed.
    pub fn from_reader_with_options<R: Read>(
        reader: R,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        options.check_owning()?;
        Self::from_opc_package_with_provenance_and_semantic_limits(
            OpcPackage::from_reader(reader, options.source_limits)?,
            true,
            options.semantic_limits,
        )
    }

    /// Parse a PPTX from a borrowed ZIP buffer.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_options(bytes, crate::OpenOptions::default())
    }

    /// Parse a PPTX from a borrowed ZIP buffer with explicit ingress options.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::UnsupportedOpenPolicy`] for cache or execution
    /// policies, and an error if the input exceeds the source limits or is
    /// malformed.
    pub fn from_bytes_with_options(bytes: &[u8], options: crate::OpenOptions) -> Result<Self> {
        options.check_owning()?;
        Self::from_opc_package_with_provenance_and_semantic_limits(
            OpcPackage::from_bytes(bytes, options.source_limits)?,
            false,
            options.semantic_limits,
        )
    }

    /// Adopt an already parsed OPC graph without hydrating a mutable model.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_opc_package(opc: OpcPackage) -> Result<Self> {
        Self::from_opc_package_with_provenance(opc, false)
    }

    pub(crate) fn from_opc_package_with_provenance(
        opc: OpcPackage,
        physical_source_provenance: bool,
    ) -> Result<Self> {
        Self::from_opc_package_with_provenance_and_semantic_limits(
            opc,
            physical_source_provenance,
            None,
        )
    }

    pub(crate) fn from_opc_package_with_provenance_and_semantic_limits(
        opc: OpcPackage,
        physical_source_provenance: bool,
        semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
    ) -> Result<Self> {
        let part_bytes = Self::part_bytes_for(&opc)?;
        let main = opc.main_document_part()?;
        let presentation = PresentationPart::from_part(main)?;
        // Slide-list repairs (audit finding 6.7) are recorded at open. A slide
        // list the catalog reader rejects stays the lazy accessors' error, as
        // before: opening never parsed the slide list to fail on it.
        let mut tolerance = crate::leniency::ToleranceReport::default();
        // OPC manifest repairs (audit findings 8.5-8.9) come first: they were
        // applied before any PresentationML part could be read.
        tolerance.append_package(opc.tolerance_report());
        let _deferred_catalog_error = presentation.slide_references_with_tolerance(&mut tolerance);
        // The document metadata reads are advisory at open, like the slide
        // list: a properties part they reject stays its accessor's error,
        // and opening never failed on it before.
        let mut properties = litchi_ooxml_common::leniency::ToleranceReport::default();
        let _deferred_core_error =
            litchi_ooxml_common::properties::read_with_report(&opc, &mut properties);
        tolerance.append_properties(&properties);
        if let Ok(custom) = Props::read_for(&opc, Host::PowerPoint) {
            tolerance.append_properties(custom.tolerance_report());
        }
        // Presentation properties (`presProps.xml`) back
        // `Package::math_properties`, a lazy read that re-reads them without
        // a report; their repairs are recorded here, advisorily, like the
        // metadata above.
        let presentation_limits =
            crate::presentation_properties::Limits::from_source_limits(opc.read_limits());
        let _deferred_presentation_properties_error =
            crate::presentation_properties::load_from_package_with_tolerance(
                &opc,
                &mut tolerance,
                presentation_limits,
            );
        let replay_index = crate::package::replay_index::ReplayIndex::from_package(&opc)?;
        Ok(Self {
            opc: std::sync::Arc::new(opc),
            part_bytes,
            content_revision: std::sync::Arc::new(std::sync::OnceLock::new()),
            part_digests: std::sync::Arc::new(std::sync::OnceLock::new()),
            replay_index: std::sync::Arc::new(replay_index),
            physical_source_provenance,
            semantic_limits,
            tolerance,
            #[cfg(feature = "encryption")]
            encryption: litchi_ooxml_common::package_encryption::PackageEncryption::plain(),
            #[cfg(feature = "encryption")]
            encryption_limits: litchi_crypto::ooxml::Limits::default(),
        })
    }

    /// Save the package atomically through the OPC writer.
    ///
    /// This is [`Self::save_with_durability`] at
    /// [`Durability::Full`](litchi_core::Durability::Full).
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        self.save_with_durability(path, litchi_core::Durability::Full)
    }

    /// Save the package atomically through the OPC writer at a caller-chosen
    /// [`Durability`](litchi_core::Durability).
    ///
    /// Every level publishes the same bytes through the same sibling
    /// temporary file and rename, and a failure before the rename leaves the
    /// destination untouched. A weaker level only skips the parent-directory
    /// sync (`FileOnly`) or both the file and directory syncs (`NoSync`),
    /// with the crash guarantees `Durability` states. The level applies to
    /// this call only and is never stored on the package.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written. Only
    /// `Durability::Full` can report that the destination was replaced but
    /// its directory could not be synchronized.
    pub fn save_with_durability<P: AsRef<Path>>(
        &self,
        path: P,
        durability: litchi_core::Durability,
    ) -> Result<()> {
        #[cfg(feature = "encryption")]
        self.encryption
            .ordinary_output()
            .map_err(|source| Error::EncryptionPolicy {
                operation: "save",
                source,
            })?;
        PackageWriter::write_with_durability(path, &self.opc, durability)?;
        Ok(())
    }

    /// Serialize the package into a new ZIP buffer.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        #[cfg(feature = "encryption")]
        self.encryption
            .ordinary_output()
            .map_err(|source| Error::EncryptionPolicy {
                operation: "to_bytes",
                source,
            })?;
        Ok(PackageWriter::to_bytes(&self.opc)?)
    }

    /// Publish an authored presentation model into this freshly created
    /// package: the slides when the model changed, and the automatic fonts
    /// when `embed_fonts` asks for them. On error the package is dropped.
    pub(crate) fn publish_authored(
        mut self,
        presentation: &crate::builder::Presentation,
        embed_fonts: bool,
        #[cfg(feature = "automatic-fonts")] preparation_limits: litchi_fonts::PreparationLimits,
        #[cfg(feature = "automatic-fonts")] eot_limits: litchi_fonts::embedding::powerpoint::Limits,
    ) -> Result<Self> {
        let _ = embed_fonts;
        let presentation_modified = presentation.is_modified();
        #[cfg(feature = "automatic-fonts")]
        let fonts_requested = self.opc.save_options().fonts != litchi_opc::FontEmbedding::None
            && (presentation_modified || embed_fonts);
        #[cfg(not(feature = "automatic-fonts"))]
        let fonts_requested = false;
        if !presentation_modified && !fonts_requested {
            return Ok(self);
        }
        if presentation_modified {
            self.materialize_presentation(presentation)?;
        }
        #[cfg(feature = "automatic-fonts")]
        self.embed_fonts_for_presentation(presentation, preparation_limits, eot_limits)?;
        self.refresh_authored_publication()?;
        Ok(self)
    }

    fn materialize_presentation(
        &mut self,
        presentation: &crate::builder::Presentation,
    ) -> Result<()> {
        // Compile all fallible p202 payloads before changing any OPC state.
        let designer = presentation.preflight_designer()?;
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
            let part = std::sync::Arc::make_mut(&mut self.opc).get_part_mut(&presentation_name)?;
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
            std::sync::Arc::make_mut(&mut self.opc).remove_part(&part);
        }
        let old_notes_slides: Vec<_> = self
            .opc
            .iter_parts()
            .filter(|part| part.partname().as_str().starts_with("/ppt/notesSlides/"))
            .map(|part| part.partname().clone())
            .collect();
        for part in old_notes_slides {
            std::sync::Arc::make_mut(&mut self.opc).remove_part(&part);
        }

        let mut relationship_ids = Vec::with_capacity(presentation.slide_count());
        for (index, (slide, prepared)) in presentation
            .slides()
            .iter()
            .zip(designer.slides())
            .enumerate()
        {
            let target = format!("slides/slide{}.xml", index.saturating_add(1));
            let relationship_id = std::sync::Arc::make_mut(&mut self.opc)
                .get_part_mut(&presentation_name)?
                .relate_to(&target, rt::SLIDE);
            relationship_ids.push(relationship_id);

            let slide_name = pack_uri(&format!("/ppt/{target}"))?;
            let mut slide_part = BlobPart::new(
                slide_name,
                ct::PML_SLIDE.to_string(),
                slide.generate_slide_xml_with(prepared)?.into_bytes(),
            );
            slide_part.relate_to("../slideLayouts/slideLayout1.xml", rt::SLIDE_LAYOUT);
            if slide.has_notes() {
                slide_part.relate_to(
                    &format!("../notesSlides/notesSlide{}.xml", index + 1),
                    rt::NOTES_SLIDE,
                );
            }
            std::sync::Arc::make_mut(&mut self.opc).add_part(Box::new(slide_part));

            if let Some(text) = slide.notes() {
                let notes_name =
                    pack_uri(&format!("/ppt/notesSlides/notesSlide{}.xml", index + 1))?;
                let mut notes_part = BlobPart::new(
                    notes_name,
                    ct::PML_NOTES_SLIDE.to_string(),
                    crate::notes::write_text(
                        text,
                        self.semantic_limits()
                            .notes()
                            .constrained_by_source(self.opc.read_limits()),
                    )?,
                );
                notes_part.relate_to(&format!("../slides/slide{}.xml", index + 1), rt::SLIDE);
                notes_part.relate_to("../notesMasters/notesMaster1.xml", rt::NOTES_MASTER);
                std::sync::Arc::make_mut(&mut self.opc).add_part(Box::new(notes_part));
            }
        }

        let xml =
            presentation.generate_presentation_xml_with_designer(&relationship_ids, &designer)?;
        std::sync::Arc::make_mut(&mut self.opc)
            .get_part_mut(&presentation_name)?
            .set_blob(xml.into_bytes());
        Ok(())
    }
}

fn pack_uri(value: &str) -> Result<PackURI> {
    PackURI::new(value).map_err(Error::Uri)
}
