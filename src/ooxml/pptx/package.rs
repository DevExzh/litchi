/// Package implementation for PowerPoint presentations.
use crate::ooxml::common::DocumentProperties;
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::OpcPackage;
use crate::ooxml::opc::constants::content_type as ct;
use crate::ooxml::opc::packuri::PackURI;
use crate::ooxml::opc::part::Part;
use crate::ooxml::pptx::parts::PresentationPart;
use crate::ooxml::pptx::presentation::Presentation;
use crate::ooxml::pptx::writer::MutablePresentation;
use std::io::{Read, Seek};
use std::path::Path;

/// A PowerPoint (.pptx) package.
///
/// This is the main entry point for working with PowerPoint presentations.
/// It wraps an OPC package and provides PowerPoint-specific functionality.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::pptx::Package;
///
/// // Open an existing presentation
/// let pkg = Package::open("presentation.pptx")?;
///
/// // Get the main presentation
/// let pres = pkg.presentation()?;
///
/// // Access slides
/// println!("Presentation has {} slides", pres.slide_count()?);
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Package {
    /// The underlying OPC package
    opc: OpcPackage,
    /// Mutable presentation for writing (cached)
    mutable_pres: Option<MutablePresentation>,
    /// Document properties (metadata)
    properties: DocumentProperties,
}

impl Package {
    /// Create a new empty .pptx package.
    ///
    /// Creates a minimal valid PowerPoint presentation with default master slide and layout.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let pkg = Package::new()?;
    /// // Add slides to the presentation...
    /// pkg.save("new_presentation.pptx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn new() -> Result<Self> {
        use crate::ooxml::opc::constants::content_type as ct;
        use crate::ooxml::opc::constants::relationship_type as rt;
        use crate::ooxml::opc::packuri::PackURI;
        use crate::ooxml::opc::part::BlobPart;
        use crate::ooxml::pptx::template;

        let mut opc = OpcPackage::new();

        // Create presentation.xml part
        let pres_partname = PackURI::new("/ppt/presentation.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("presentation partname: {}", e)))?;
        let pres_part = BlobPart::new(
            pres_partname.clone(),
            ct::PML_PRESENTATION_MAIN.to_string(),
            template::default_presentation_xml().as_bytes().to_vec(),
        );

        // Create relationship from package to presentation (use relative path for package-level rels)
        opc.relate_to("ppt/presentation.xml", rt::OFFICE_DOCUMENT);
        opc.add_part(Box::new(pres_part));

        // Create slideMaster.xml
        let master_partname = PackURI::new("/ppt/slideMasters/slideMaster1.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("slideMaster partname: {}", e)))?;
        let master_part = BlobPart::new(
            master_partname.clone(),
            ct::PML_SLIDE_MASTER.to_string(),
            template::default_slide_master_xml().as_bytes().to_vec(),
        );

        // Add relationship from presentation to slideMaster
        if let Ok(pres_part) = opc.get_part_mut(&pres_partname) {
            pres_part.relate_to("slideMasters/slideMaster1.xml", rt::SLIDE_MASTER);
        }
        opc.add_part(Box::new(master_part));

        // Create all 11 slide layouts
        // Each layout needs:
        // 1. A relationship FROM slideMaster TO the layout
        // 2. A relationship FROM the layout back TO slideMaster
        let layout_xmls = template::all_slide_layouts();
        for (i, layout_xml) in layout_xmls.iter().enumerate() {
            let layout_num = i + 1;
            let layout_partname_str = format!("/ppt/slideLayouts/slideLayout{}.xml", layout_num);
            let layout_partname = PackURI::new(&layout_partname_str).map_err(|e| {
                OoxmlError::InvalidUri(format!("slideLayout{} partname: {}", layout_num, e))
            })?;

            let mut layout_part = BlobPart::new(
                layout_partname.clone(),
                ct::PML_SLIDE_LAYOUT.to_string(),
                layout_xml.as_bytes().to_vec(),
            );

            // Add relationship from slideMaster to this slideLayout
            if let Ok(master_part) = opc.get_part_mut(&master_partname) {
                let layout_rel_target = format!("../slideLayouts/slideLayout{}.xml", layout_num);
                master_part.relate_to(&layout_rel_target, rt::SLIDE_LAYOUT);
            }

            // Add relationship from slideLayout back to slideMaster
            // This bidirectional relationship is required by PowerPoint
            layout_part.relate_to("../slideMasters/slideMaster1.xml", rt::SLIDE_MASTER);

            opc.add_part(Box::new(layout_part));
        }

        // Create theme.xml
        let theme_partname = PackURI::new("/ppt/theme/theme1.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("theme partname: {}", e)))?;
        let theme_part = BlobPart::new(
            theme_partname,
            ct::OFC_THEME.to_string(),
            template::default_theme_xml().as_bytes().to_vec(),
        );

        // Add relationship from slideMaster to theme
        if let Ok(master_part) = opc.get_part_mut(&master_partname) {
            master_part.relate_to("../theme/theme1.xml", rt::THEME);
        }
        opc.add_part(Box::new(theme_part));

        // Create tableStyles.xml
        let table_styles_partname = PackURI::new("/ppt/tableStyles.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("tableStyles partname: {}", e)))?;
        let table_styles_part = BlobPart::new(
            table_styles_partname,
            ct::PML_TABLE_STYLES.to_string(),
            template::default_table_styles_xml().as_bytes().to_vec(),
        );

        // Add relationship from presentation to tableStyles
        if let Ok(pres_part) = opc.get_part_mut(&pres_partname) {
            pres_part.relate_to("tableStyles.xml", rt::TABLE_STYLES);
        }
        opc.add_part(Box::new(table_styles_part));

        // Create viewProps.xml
        let view_props_partname = PackURI::new("/ppt/viewProps.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("viewProps partname: {}", e)))?;
        let view_props_part = BlobPart::new(
            view_props_partname,
            ct::PML_VIEW_PROPS.to_string(),
            template::default_view_props_xml().as_bytes().to_vec(),
        );

        // Add relationship from presentation to viewProps
        if let Ok(pres_part) = opc.get_part_mut(&pres_partname) {
            pres_part.relate_to("viewProps.xml", rt::VIEW_PROPS);
        }
        opc.add_part(Box::new(view_props_part));

        // Create presProps.xml
        let pres_props_partname = PackURI::new("/ppt/presProps.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("presProps partname: {}", e)))?;
        let pres_props_part = BlobPart::new(
            pres_props_partname,
            ct::PML_PRES_PROPS.to_string(),
            template::default_pres_props_xml().as_bytes().to_vec(),
        );

        // Add relationship from presentation to presProps
        if let Ok(pres_part) = opc.get_part_mut(&pres_partname) {
            pres_part.relate_to("presProps.xml", rt::PRES_PROPS);
        }
        opc.add_part(Box::new(pres_props_part));

        // Create core.xml (core properties)
        let core_props_partname = PackURI::new("/docProps/core.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("core.xml partname: {}", e)))?;
        let core_props_part = BlobPart::new(
            core_props_partname,
            ct::OPC_CORE_PROPERTIES.to_string(),
            template::default_core_props_xml().as_bytes().to_vec(),
        );

        opc.relate_to("docProps/core.xml", rt::CORE_PROPERTIES);
        opc.add_part(Box::new(core_props_part));

        // Create app.xml (extended properties)
        let app_props_partname = PackURI::new("/docProps/app.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("app.xml partname: {}", e)))?;
        let app_props_part = BlobPart::new(
            app_props_partname,
            ct::OFC_EXTENDED_PROPERTIES.to_string(),
            template::default_app_props_xml().as_bytes().to_vec(),
        );

        opc.relate_to("docProps/app.xml", rt::EXTENDED_PROPERTIES);
        opc.add_part(Box::new(app_props_part));

        // Create a mutable presentation for writing
        let mutable_pres = Some(MutablePresentation::new());

        // Initialize document properties
        let properties = DocumentProperties::new();

        Ok(Self {
            opc,
            mutable_pres,
            properties,
        })
    }

    /// Open a .pptx package from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .pptx file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let opc = OpcPackage::open(path)?;

        // Verify it's a PowerPoint presentation by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main presentation part: {}", e)))?;

        let content_type = main_part.content_type();
        // Support both regular and macro-enabled presentations
        if content_type != ct::PML_PRESENTATION_MAIN && content_type != ct::PML_PRES_MACRO_MAIN {
            return Err(OoxmlError::InvalidContentType {
                expected: format!(
                    "{} or {}",
                    ct::PML_PRESENTATION_MAIN,
                    ct::PML_PRES_MACRO_MAIN
                ),
                got: content_type.to_string(),
            });
        }

        Ok(Self {
            opc,
            mutable_pres: None,
            properties: DocumentProperties::new(),
        })
    }

    /// Create a Package from an already-parsed OPC package.
    ///
    /// This is used for single-pass parsing where the OPC package has already
    /// been parsed during format detection. It avoids double-parsing.
    ///
    /// # Arguments
    ///
    /// * `opc` - An already-parsed OPC package
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::{OpcPackage, pptx::Package};
    /// use std::io::Cursor;
    ///
    /// let bytes = std::fs::read("presentation.pptx")?;
    /// let opc = OpcPackage::from_reader(Cursor::new(bytes))?;
    /// let pkg = Package::from_opc_package(opc)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_opc_package(opc: OpcPackage) -> Result<Self> {
        // Verify it's a PowerPoint presentation by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main presentation part: {}", e)))?;

        let content_type = main_part.content_type();
        // Support both regular and macro-enabled presentations
        if content_type != ct::PML_PRESENTATION_MAIN && content_type != ct::PML_PRES_MACRO_MAIN {
            return Err(OoxmlError::InvalidContentType {
                expected: format!(
                    "{} or {}",
                    ct::PML_PRESENTATION_MAIN,
                    ct::PML_PRES_MACRO_MAIN
                ),
                got: content_type.to_string(),
            });
        }

        Ok(Self {
            opc,
            mutable_pres: None,
            properties: DocumentProperties::new(),
        })
    }

    /// Create a .pptx package from a reader.
    ///
    /// # Arguments
    ///
    /// * `reader` - A reader containing the .pptx file data (must implement Read + Seek)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    /// use std::io::Cursor;
    ///
    /// let data = std::fs::read("presentation.pptx")?;
    /// let cursor = Cursor::new(data);
    /// let pkg = Package::from_reader(cursor)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let opc = OpcPackage::from_reader(reader)?;

        // Verify it's a PowerPoint presentation by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main presentation part: {}", e)))?;

        let content_type = main_part.content_type();
        // Support both regular and macro-enabled presentations
        if content_type != ct::PML_PRESENTATION_MAIN && content_type != ct::PML_PRES_MACRO_MAIN {
            return Err(OoxmlError::InvalidContentType {
                expected: format!(
                    "{} or {}",
                    ct::PML_PRESENTATION_MAIN,
                    ct::PML_PRES_MACRO_MAIN
                ),
                got: content_type.to_string(),
            });
        }

        Ok(Self {
            opc,
            mutable_pres: None,
            properties: DocumentProperties::new(),
        })
    }

    /// Get the main presentation for reading.
    ///
    /// Returns the `Presentation` object which provides access to the presentation's
    /// content, slides, and other features.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// // Access slides
    /// for slide in pres.slides()? {
    ///     println!("Slide text: {}", slide.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn presentation(&self) -> Result<Presentation<'_>> {
        let main_part = self
            .opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main presentation part: {}", e)))?;

        // Create PresentationPart wrapper
        let pres_part = PresentationPart::from_part(main_part)?;

        // Create and return Presentation
        Ok(Presentation::new(pres_part, &self.opc))
    }

    /// Get the underlying OPC package.
    ///
    /// This provides access to lower-level package operations.
    #[inline]
    pub fn opc_package(&self) -> &OpcPackage {
        &self.opc
    }

    /// Get mutable access to the underlying OPC package.
    ///
    /// This provides access to lower-level package operations for modification.
    #[inline]
    pub fn opc_package_mut(&mut self) -> &mut OpcPackage {
        &mut self.opc
    }

    /// Get a mutable presentation for writing and modification.
    ///
    /// This returns a `MutablePresentation` that allows you to add and modify
    /// slides, shapes, and other presentation elements.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// let mut pres = pkg.presentation_mut()?;
    ///
    /// // Add a slide
    /// let slide = pres.add_slide()?;
    /// slide.set_title("My Presentation");
    /// slide.add_text_box("Hello, World!", 914400, 914400, 2743200, 914400);
    ///
    /// pkg.save("output.pptx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn presentation_mut(&mut self) -> Result<&mut MutablePresentation> {
        // If we don't have a mutable presentation, create one
        if self.mutable_pres.is_none() {
            self.mutable_pres = Some(MutablePresentation::new());
        }

        Ok(self.mutable_pres.as_mut().unwrap())
    }

    /// Get a reference to the presentation properties.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let props = pkg.properties();
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn properties(&self) -> &DocumentProperties {
        &self.properties
    }

    /// Get a mutable reference to the presentation properties.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// pkg.properties_mut().title = Some("My Presentation".to_string());
    /// pkg.properties_mut().creator = Some("John Doe".to_string());
    /// pkg.save("presentation.pptx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn properties_mut(&mut self) -> &mut DocumentProperties {
        &mut self.properties
    }

    /// Save the package to a file.
    ///
    /// Writes the complete PowerPoint presentation including all parts, relationships,
    /// and content types to a .pptx file.
    ///
    /// # Arguments
    /// * `path` - Path where the .pptx file should be written
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// // Modify presentation...
    /// pkg.save("output.pptx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        // If we have a mutable presentation, update the presentation parts
        let should_update = self
            .mutable_pres
            .as_ref()
            .map(|p| p.is_modified())
            .unwrap_or(false);

        if should_update {
            // Take mutable_pres temporarily to avoid borrow issues
            if let Some(mutable_pres) = self.mutable_pres.take() {
                self.update_presentation_parts(&mutable_pres)?;
                self.mutable_pres = Some(mutable_pres);
            }
        }

        // Update core properties
        self.update_core_properties()?;

        self.opc.save(path).map_err(|e| {
            OoxmlError::IoError(std::io::Error::other(format!(
                "Failed to save package: {}",
                e
            )))
        })
    }

    /// Update presentation parts with modified data.
    fn update_presentation_parts(&mut self, pres: &MutablePresentation) -> Result<()> {
        use crate::ooxml::opc::constants::content_type as ct;
        use crate::ooxml::opc::constants::relationship_type as rt;
        use crate::ooxml::opc::part::{BlobPart, Part};
        use crate::ooxml::pptx::writer::relmap::RelationshipMapper;

        // Initialize relationship mapper
        let mut rel_mapper = RelationshipMapper::new();

        // Collect all images from all slides
        let all_images = pres.collect_all_images();

        // Create image parts first and add to package
        for (img_index, (_slide_index, image_data, image_format)) in all_images.iter().enumerate() {
            let img_num = img_index + 1;
            let ext = image_format.extension();

            // Create image part URI
            let image_partname = format!("/ppt/media/image{}.{}", img_num, ext);
            let image_uri = PackURI::new(&image_partname)
                .map_err(|e| OoxmlError::InvalidUri(format!("image URI: {}", e)))?;

            // Create image part
            let image_part = BlobPart::new(
                image_uri,
                image_format.mime_type().to_string(),
                image_data.to_vec(),
            );

            // Add image part to package
            self.opc.add_part(Box::new(image_part));
        }

        // Create presentation part and add relationships
        let pres_uri = PackURI::new("/ppt/presentation.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("presentation URI: {}", e)))?;

        // Create a temporary presentation part to manage relationships
        let mut temp_pres_part = BlobPart::new(
            pres_uri.clone(),
            ct::PML_PRESENTATION_MAIN.to_string(),
            Vec::new(),
        );

        // Add relationship to slideMaster (this should be rId1)
        temp_pres_part.relate_to("slideMasters/slideMaster1.xml", rt::SLIDE_MASTER);

        // Add other required relationships (in the order they were created in Package::new())
        // These relationships should be added even if not modified, as they're required for a valid PPTX
        temp_pres_part.relate_to("tableStyles.xml", rt::TABLE_STYLES);
        temp_pres_part.relate_to("viewProps.xml", rt::VIEW_PROPS);
        temp_pres_part.relate_to("presProps.xml", rt::PRES_PROPS);

        // Track slide relationship IDs for presentation.xml generation
        let mut slide_rel_ids: Vec<String> = Vec::new();

        // Process each slide: create relationships first, then generate XML
        for (slide_index, slide) in pres.slides.iter().enumerate() {
            if slide.is_modified() {
                let slide_num = slide_index + 1;
                let slide_uri = PackURI::new(format!("/ppt/slides/slide{}.xml", slide_num))
                    .map_err(|e| {
                        OoxmlError::InvalidUri(format!("slide{} URI: {}", slide_num, e))
                    })?;

                // Create a temporary slide part to manage relationships
                let mut temp_slide_part =
                    BlobPart::new(slide_uri.clone(), ct::PML_SLIDE.to_string(), Vec::new());

                // Add relationship from slide to slide layout (always first relationship)
                temp_slide_part.relate_to("../slideLayouts/slideLayout1.xml", rt::SLIDE_LAYOUT);

                // Collect images for this slide and create relationships
                let slide_images = slide.collect_images();
                for (img_index_in_slide, (_, image_format)) in slide_images.iter().enumerate() {
                    // Find the global image index for this slide's image
                    let mut global_img_idx = 0;
                    for (global_idx, (s_idx, _, _)) in all_images.iter().enumerate() {
                        if *s_idx == slide_index {
                            if global_img_idx == img_index_in_slide {
                                let img_num = global_idx + 1;
                                let ext = image_format.extension();
                                let image_rel_target = format!("../media/image{}.{}", img_num, ext);
                                let rid = temp_slide_part.relate_to(&image_rel_target, rt::IMAGE);
                                rel_mapper.add_image(slide_index, img_index_in_slide, rid);
                                break;
                            }
                            global_img_idx += 1;
                        }
                    }
                }

                // Add relationship from slide to notes slide if notes exist
                if slide.has_notes() {
                    let notes_rel_target = format!("../notesSlides/notesSlide{}.xml", slide_num);
                    let rid = temp_slide_part.relate_to(&notes_rel_target, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide");
                    rel_mapper.add_notes(slide_index, rid);
                }

                // Now generate slide XML with actual relationship IDs
                let slide_xml = slide.to_xml_with_rels(Some(slide_index), Some(&rel_mapper))?;

                // Update the temp part with the actual XML content
                temp_slide_part.set_blob(slide_xml.into_bytes());

                // Add the slide part to the package
                self.opc.add_part(Box::new(temp_slide_part));

                // Create notes slide if notes exist
                if let Some(notes_xml_result) = slide.generate_notes_xml() {
                    let notes_xml = notes_xml_result?;
                    let notes_uri =
                        PackURI::new(format!("/ppt/notesSlides/notesSlide{}.xml", slide_num))
                            .map_err(|e| {
                                OoxmlError::InvalidUri(format!(
                                    "notesSlide{} URI: {}",
                                    slide_num, e
                                ))
                            })?;

                    let mut notes_part = BlobPart::new(
                        notes_uri,
                        "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml".to_string(),
                        notes_xml.into_bytes(),
                    );

                    // Add relationship from notes to slide
                    notes_part.relate_to(&format!("../slides/slide{}.xml", slide_num), rt::SLIDE);

                    self.opc.add_part(Box::new(notes_part));
                }

                // Add relationship from presentation to this slide and track the ID
                let rel_target = format!("slides/slide{}.xml", slide_num);
                let slide_rid = temp_pres_part.relate_to(&rel_target, rt::SLIDE);
                slide_rel_ids.push(slide_rid);
            }
        }

        // Now generate presentation XML with actual relationship IDs
        let pres_xml = pres.generate_presentation_xml_with_rels(Some(&slide_rel_ids))?;
        temp_pres_part.set_blob(pres_xml.into_bytes());

        // Add the presentation part to the package
        self.opc.add_part(Box::new(temp_pres_part));

        Ok(())
    }

    /// Update the core.xml properties part.
    fn update_core_properties(&mut self) -> Result<()> {
        use crate::ooxml::opc::part::BlobPart;

        let core_uri = PackURI::new("/docProps/core.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("core.xml URI: {}", e)))?;

        // Generate XML from properties
        let xml = self.properties.to_xml();

        // Create or update the core properties part
        let core_part = BlobPart::new(
            core_uri,
            ct::OPC_CORE_PROPERTIES.to_string(),
            xml.into_bytes(),
        );

        self.opc.add_part(Box::new(core_part));

        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    #[ignore] // Requires test file
    fn test_open_package() {
        let result = Package::open("test.pptx");
        assert!(result.is_ok());
    }
}
