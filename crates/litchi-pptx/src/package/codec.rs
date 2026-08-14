//! OPC package graph construction, publication, and serialization.

use std::io::Read;
use std::path::Path;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::packuri::PackURI;
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackageWriter};

use super::Package;
use crate::parts::PresentationPart;
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

pub(super) const STALE_PRESENTATION_GRAPH_REASON: &str = "the mutable presentation model has unflushed changes; save and reopen before reading the canonical package graph";

impl Package {
    /// Create a minimal valid, mutable `PresentationML` package.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
            physical_source_provenance: false,
            #[cfg(feature = "encryption")]
            encryption: litchi_ooxml_common::package_encryption::PackageEncryption::plain(),
            #[cfg(feature = "automatic-fonts")]
            font_embedding_dirty: false,
        })
    }

    /// Open a PPTX from a filesystem path.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::open_with_limits(path, litchi_opc::ReadLimits::default())
    }

    /// Open a PPTX from a filesystem path with explicit OPC resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn open_with_limits<P: AsRef<Path>>(
        path: P,
        limits: litchi_opc::ReadLimits,
    ) -> Result<Self> {
        Self::from_opc_package_with_provenance(OpcPackage::open_with_limits(path, limits)?, true)
    }

    /// Parse a PPTX from a reader.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_reader<R: Read>(reader: R) -> Result<Self> {
        Self::from_reader_with_limits(reader, litchi_opc::ReadLimits::default())
    }

    /// Parse a PPTX from a reader with explicit OPC resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_reader_with_limits<R: Read>(
        reader: R,
        limits: litchi_opc::ReadLimits,
    ) -> Result<Self> {
        Self::from_opc_package_with_provenance(
            OpcPackage::from_reader_with_limits(reader, limits)?,
            true,
        )
    }

    /// Parse a PPTX from an owned ZIP buffer.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_vec(bytes: Vec<u8>) -> Result<Self> {
        Self::from_vec_with_limits(bytes, litchi_opc::ReadLimits::default())
    }

    /// Parse a PPTX from an owned ZIP buffer with explicit OPC resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_vec_with_limits(bytes: Vec<u8>, limits: litchi_opc::ReadLimits) -> Result<Self> {
        Self::from_opc_package_with_provenance(
            OpcPackage::from_vec_with_limits(bytes, limits)?,
            true,
        )
    }

    /// Parse a PPTX from a borrowed ZIP buffer.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, litchi_opc::ReadLimits::default())
    }

    /// Parse a PPTX from a borrowed ZIP buffer with explicit OPC resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: litchi_opc::ReadLimits) -> Result<Self> {
        Self::from_opc_package(OpcPackage::from_bytes_with_limits(bytes, limits)?)
    }

    /// Adopt an already parsed OPC graph without hydrating a mutable model.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_opc_package(opc: OpcPackage) -> Result<Self> {
        Self::from_opc_package_with_provenance(opc, false)
    }

    fn from_opc_package_with_provenance(
        opc: OpcPackage,
        physical_source_provenance: bool,
    ) -> Result<Self> {
        let main = opc.main_document_part()?;
        PresentationPart::from_part(main)?;
        Ok(Self {
            opc,
            mutable_pres: None,
            physical_source_provenance,
            #[cfg(feature = "encryption")]
            encryption: litchi_ooxml_common::package_encryption::PackageEncryption::plain(),
            #[cfg(feature = "automatic-fonts")]
            font_embedding_dirty: false,
        })
    }

    /// Save the package atomically through the OPC writer.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        #[cfg(feature = "encryption")]
        self.encryption
            .ordinary_output()
            .map_err(|source| Error::EncryptionPolicy {
                operation: "save",
                source,
            })?;
        self.flush_presentation()?;
        PackageWriter::write(path, &self.opc)?;
        Ok(())
    }

    /// Serialize the package into a new ZIP buffer.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_bytes(&mut self) -> Result<Vec<u8>> {
        #[cfg(feature = "encryption")]
        self.encryption
            .ordinary_output()
            .map_err(|source| Error::EncryptionPolicy {
                operation: "to_bytes",
                source,
            })?;
        self.flush_presentation()?;
        Ok(PackageWriter::to_bytes(&self.opc)?)
    }

    pub(super) fn edit_typed<T>(
        &mut self,
        operation: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        self.ensure_plain_mutation("typed package mutation")?;
        self.edit_raw(operation)
    }

    pub(super) fn edit_raw<T>(
        &mut self,
        operation: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        let before = self.opc.clone();
        let presentation_before = self.mutable_pres.clone();
        #[cfg(feature = "automatic-fonts")]
        let font_embedding_dirty_before = self.font_embedding_dirty;

        let result = (|| {
            self.flush_presentation()?;
            operation(&mut self.opc)
        })();
        match result {
            Ok(value) => Ok(value),
            Err(error) => {
                self.opc = before;
                self.mutable_pres = presentation_before;
                #[cfg(feature = "automatic-fonts")]
                {
                    self.font_embedding_dirty = font_embedding_dirty_before;
                }
                Err(error)
            },
        }
    }

    pub(super) fn flush_presentation(&mut self) -> Result<()> {
        let presentation_modified = self.is_modified();
        #[cfg(feature = "automatic-fonts")]
        let fonts_requested = self.opc.save_options().fonts != litchi_opc::FontEmbedding::None
            && (presentation_modified || self.font_embedding_dirty);
        #[cfg(not(feature = "automatic-fonts"))]
        let fonts_requested = false;
        if !presentation_modified && !fonts_requested {
            return Ok(());
        }
        let Some(presentation) = self.mutable_pres.clone() else {
            return Err(Error::UnsafeEdit {
                operation: "save",
                reason: "the lossless facade cannot publish an opened package's mutable graph",
            });
        };
        let before = self.opc.clone();
        let result = (|| {
            if presentation_modified {
                self.materialize_presentation(&presentation)?;
            }
            #[cfg(feature = "automatic-fonts")]
            self.embed_fonts_for_presentation(&presentation)?;
            Ok(())
        })();
        match result {
            Ok(()) => {
                #[cfg(feature = "automatic-fonts")]
                if fonts_requested {
                    self.font_embedding_dirty = false;
                }
                if presentation_modified && let Some(presentation) = self.mutable_pres.as_mut() {
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
        for (index, (slide, prepared)) in presentation
            .slides()
            .iter()
            .zip(designer.slides())
            .enumerate()
        {
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
                slide.generate_slide_xml_with(prepared)?.into_bytes(),
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

        let xml =
            presentation.generate_presentation_xml_with_designer(&relationship_ids, &designer)?;
        self.opc
            .get_part_mut(&presentation_name)?
            .set_blob(xml.into_bytes());
        Ok(())
    }
}

fn pack_uri(value: &str) -> Result<PackURI> {
    PackURI::new(value).map_err(Error::Uri)
}
